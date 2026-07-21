from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from threading import Thread
from typing import Any
from uuid import uuid4

from flask import Blueprint, current_app, jsonify, request, url_for

from ..encryption import decrypt_file, encrypt_file, is_encrypted
from ..excel_io import load_workbook_with_warnings
from ..merge_engine import MergeEngine
from ..merge_models import MergeJob, MergeSheetConfig, MergeSummary
from ..merge_planning import MergePlan, build_merge_plan
from .routes import _jobs, _json_payload, _json_value, _required_text


merge_api = Blueprint("excel_splitter_merge", __name__, url_prefix="/api/merge")


@merge_api.post("/load")
def load_files():
    uploads = request_files()
    if not uploads:
        raise ValueError("请选择 Excel 文件")

    job_id = request_form_text("job_id")
    if job_id:
        job = _get_merge_job(job_id)
    else:
        job_id = uuid4().hex
        job = {"kind": "merge", "files": [], "results": []}
        _jobs()[job_id] = job

    source_paths = request_form_list("source_paths")
    job_dir = current_app.config["UPLOAD_DIR"] / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    for index, upload in enumerate(uploads):
        original_filename = Path(upload.filename or "").name
        if not original_filename:
            raise ValueError("请选择 Excel 文件")
        if Path(original_filename).suffix.lower() != ".xlsx":
            raise ValueError("第一版仅支持 .xlsx 文件")
        if any(item["filename"] == original_filename for item in job["files"]):
            raise ValueError(f"已存在同名文件：{original_filename}")
        file_id = uuid4().hex
        # 以原始文件名存放：合并报告与「来源文件」列按文件名展示
        stored_path = job_dir / original_filename
        upload.save(stored_path)
        job["files"].append(
            {
                "file_id": file_id,
                "filename": original_filename,
                "stored_path": stored_path,
                "size": stored_path.stat().st_size,
                "encrypted": is_encrypted(stored_path),
                "source_path": source_paths[index] if index < len(source_paths) else "",
            }
        )
    _invalidate_plan(job)
    return jsonify(
        job_id=job_id,
        files=[_serialize_file(item) for item in job["files"]],
        has_encrypted=any(item["encrypted"] for item in job["files"]),
    )


@merge_api.post("/unlock")
def unlock_files():
    payload = _json_payload()
    job = _get_merge_job(payload.get("job_id"))
    unified_password = payload.get("password")
    if not isinstance(unified_password, str):
        unified_password = ""
    file_passwords = payload.get("file_passwords")
    if not isinstance(file_passwords, dict):
        file_passwords = {}

    results: list[dict[str, Any]] = []
    for item in job["files"]:
        if not item["encrypted"]:
            continue
        password = file_passwords.get(item["file_id"]) or unified_password
        if not password:
            results.append(_unlock_result(item, False, "缺少密码"))
            continue
        stored_path = Path(item["stored_path"])
        try:
            decrypted = decrypt_file(stored_path, password)
        except Exception:
            results.append(_unlock_result(item, False, "密码错误，无法解密文件"))
            continue
        # 解密文件放入 decrypted 子目录并保留原始文件名，
        # 保证 path.name 仍是用户文件名（合并报告与来源列按文件名展示）
        decrypted_dir = stored_path.parent / "decrypted"
        decrypted_dir.mkdir(exist_ok=True)
        decrypted_path = decrypted_dir / stored_path.name
        with open(decrypted_path, "wb") as target:
            target.write(decrypted.read())
        try:
            workbook, _ = load_workbook_with_warnings(
                decrypted_path, data_only=True, read_only=True
            )
            workbook.close()
        except Exception:
            decrypted_path.unlink(missing_ok=True)
            results.append(_unlock_result(item, False, "解密后文件无法读取"))
            continue
        item["stored_path"] = decrypted_path
        item["encrypted"] = False
        results.append(_unlock_result(item, True, None))
    _invalidate_plan(job)
    return jsonify(
        results=results,
        has_encrypted=any(item["encrypted"] for item in job["files"]),
    )


@merge_api.post("/files/remove")
def remove_file():
    payload = _json_payload()
    job = _get_merge_job(payload.get("job_id"))
    file_id = _required_text(payload, "file_id")
    target = next(
        (item for item in job["files"] if item["file_id"] == file_id), None
    )
    if target is None:
        raise ValueError("文件不存在或已移除")
    job["files"].remove(target)
    Path(target["stored_path"]).unlink(missing_ok=True)
    _invalidate_plan(job)
    return jsonify(
        files=[_serialize_file(item) for item in job["files"]],
        has_encrypted=any(item["encrypted"] for item in job["files"]),
    )


@merge_api.post("/files/reorder")
def reorder_files():
    payload = _json_payload()
    job = _get_merge_job(payload.get("job_id"))
    file_ids = payload.get("file_ids")
    if not isinstance(file_ids, list) or not all(
        isinstance(item, str) for item in file_ids
    ):
        raise ValueError("file_ids 必须是文件 id 列表")
    by_id = {item["file_id"]: item for item in job["files"]}
    if set(file_ids) != set(by_id) or len(file_ids) != len(by_id):
        raise ValueError("file_ids 必须与当前文件列表一致")
    job["files"] = [by_id[file_id] for file_id in file_ids]
    _invalidate_plan(job)
    return jsonify(files=[_serialize_file(item) for item in job["files"]])


@merge_api.post("/sheets")
def list_sheets():
    payload = _json_payload()
    job = _get_merge_job(payload.get("job_id"))
    sheets: list[str] = []
    for item in job["files"]:
        if item["encrypted"]:
            continue  # 未解密的文件无法读取，跳过
        workbook, _ = load_workbook_with_warnings(
            Path(item["stored_path"]), data_only=True, read_only=True
        )
        try:
            for sheet_name in workbook.sheetnames:
                if sheet_name not in sheets:
                    sheets.append(sheet_name)
        finally:
            workbook.close()
    return jsonify(sheets=sheets)


@merge_api.post("/preview")
def merge_preview():
    """预览某个 sheet 的前若干行（取自第一个包含该 sheet 的已解密文件）。"""
    payload = _json_payload()
    job = _get_merge_job(payload.get("job_id"))
    sheet_name = _required_text(payload, "sheet_name")
    try:
        max_rows = max(1, int(payload.get("max_rows", 50) or 50))
    except (TypeError, ValueError):
        raise ValueError("max_rows 必须是数字")
    for item in job["files"]:
        if item["encrypted"]:
            continue  # 未解密的文件无法读取，跳过
        workbook, _ = load_workbook_with_warnings(
            Path(item["stored_path"]), data_only=True, read_only=True
        )
        try:
            if sheet_name not in workbook.sheetnames:
                continue
            sheet = workbook[sheet_name]
            rows = [
                [_json_value(value) for value in row]
                for row in sheet.iter_rows(min_row=1, max_row=max_rows, values_only=True)
            ]
            return jsonify(
                sheet_name=sheet_name,
                source_file=item["filename"],
                rows=rows,
                total_rows=sheet.max_row or len(rows),
                max_rows=max_rows,
            )
        finally:
            workbook.close()
    raise ValueError(f"没有文件包含 sheet：{sheet_name}")


@merge_api.post("/plan")
def merge_plan():
    payload = _json_payload()
    job = _get_merge_job(payload.get("job_id"))
    configs = _merge_sheet_configs(payload.get("sheet_configs"))
    # identical 的 sheet 用户已确认内容一致，跳过字段扫描，返回标记即可
    identical_names = [config.sheet_name for config in configs if config.identical]
    scan_configs = tuple(config for config in configs if not config.identical)
    plan = (
        build_merge_plan(_readable_files(job), scan_configs)
        if scan_configs
        else MergePlan(sheets=[])
    )
    job["_merge_plan"] = plan
    job["_merge_plan_configs"] = configs
    result = _serialize_plan(plan)
    result["sheets"] = [
        {"sheet_name": name, "identical": True} for name in identical_names
    ] + result["sheets"]
    return jsonify(result)


@merge_api.post("/execute")
def execute_merge():
    payload = _json_payload()
    job_record = _get_merge_job(payload.get("job_id"))
    job_id = payload["job_id"]
    configs = _merge_sheet_configs(payload.get("sheet_configs"))
    files = job_record["files"]
    if len(files) < 2:
        raise ValueError("至少需要两个输入文件")
    locked = [item for item in files if item["encrypted"]]
    if locked:
        raise ValueError(
            f"文件 {locked[0]['filename']} 尚未解密，请先输入密码或移除该文件"
        )

    output_dir = _resolve_output_dir(payload, files[0])
    output_filename = payload.get("output_filename")
    if not isinstance(output_filename, str) or not output_filename.strip():
        output_filename = f"{Path(files[0]['filename']).stem}_合并结果.xlsx"
    source_column_name = payload.get("source_column_name")
    if not isinstance(source_column_name, str) or not source_column_name.strip():
        source_column_name = "来源文件"

    merge_job = MergeJob(
        input_files=tuple(Path(item["stored_path"]) for item in files),
        output_file=output_dir / output_filename.strip(),
        sheet_configs=configs,
        include_source_column=bool(payload.get("include_source_column", False)),
        source_column_name=source_column_name.strip(),
        skip_duplicate_sheets=bool(payload.get("skip_duplicate_sheets", True)),
        overwrite=bool(payload.get("overwrite", False)),
    )
    merge_job.validate()
    output_password = payload.get("output_password", "") or ""
    cached_plan = (
        job_record.get("_merge_plan")
        if job_record.get("_merge_plan_configs") == configs
        else None
    )

    if bool(payload.get("background", True)):
        execution = job_record.get("execution")
        if execution and execution.get("status") in ("queued", "running"):
            raise ValueError("当前任务正在执行，请勿重复提交")
        job_record["execution"] = {
            "status": "queued",
            "progress": 0,
            "message": "任务已排队",
            "summary": None,
            "error": None,
        }
        Thread(
            target=_run_background_merge,
            args=(job_record, merge_job, output_password, cached_plan),
            daemon=True,
        ).start()
        return (
            jsonify(
                status="queued",
                progress_url=url_for(
                    "excel_splitter.execution_progress", job_id=job_id
                ),
            ),
            202,
        )

    summary = MergeEngine().execute(merge_job, plan=cached_plan)
    if output_password:
        encrypt_file(summary.output_file, output_password)
    job_record["results"] = [summary.output_file]
    return jsonify(_serialize_merge_summary(summary))


def _run_background_merge(
    job_record: dict[str, Any],
    merge_job: MergeJob,
    output_password: str = "",
    plan: MergePlan | None = None,
) -> None:
    execution = job_record["execution"]

    def update_progress(percent: int, message: str) -> None:
        execution.update(
            status="running",
            progress=percent,
            message=message,
        )

    try:
        summary = MergeEngine().execute(
            merge_job, progress_callback=update_progress, plan=plan
        )
        if output_password:
            encrypt_file(summary.output_file, output_password)
        job_record["results"] = [summary.output_file]
        execution.update(
            status="complete",
            progress=100,
            message="合并完成",
            summary=_to_progress_summary(summary),
        )
    except Exception as exc:
        execution.update(
            status="failed",
            message="合并失败",
            error=str(exc),
        )


def _get_merge_job(job_id: Any) -> dict[str, Any]:
    if not isinstance(job_id, str) or job_id not in _jobs():
        raise ValueError("任务不存在或已过期，请重新选择文件")
    job = _jobs()[job_id]
    if job.get("kind") != "merge":
        raise ValueError("任务不存在或已过期，请重新选择文件")
    return job


def _readable_files(job: dict[str, Any]) -> tuple[Path, ...]:
    return tuple(
        Path(item["stored_path"]) for item in job["files"] if not item["encrypted"]
    )


def _invalidate_plan(job: dict[str, Any]) -> None:
    job["_merge_plan"] = None
    job["_merge_plan_configs"] = None


def _resolve_output_dir(payload: dict[str, Any], first_file: dict[str, Any]) -> Path:
    raw = payload.get("output_dir")
    if isinstance(raw, str) and raw.strip():
        output_dir = Path(raw.strip())
        if not output_dir.is_absolute():
            raise ValueError("输出目录必须是本机绝对路径")
        return output_dir
    source_path = first_file.get("source_path") or ""
    if source_path and Path(source_path).parent.is_dir():
        return Path(source_path).parent
    return current_app.config["DEFAULT_OUTPUT_DIR"]


def _merge_sheet_configs(raw_configs: Any) -> tuple[MergeSheetConfig, ...]:
    if not isinstance(raw_configs, list) or not raw_configs:
        raise ValueError("至少配置一个 sheet")
    configs: list[MergeSheetConfig] = []
    for raw in raw_configs:
        if not isinstance(raw, dict):
            raise ValueError("sheet 配置格式无效")
        try:
            config = MergeSheetConfig(
                sheet_name=str(raw["sheet_name"]),
                header_row=int(raw.get("header_row", 1)),
                identical=bool(raw.get("identical", False)),
            )
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError("sheet 配置缺少有效的表头行") from exc
        config.validate()
        configs.append(config)
    return tuple(configs)


def _serialize_file(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "file_id": item["file_id"],
        "filename": item["filename"],
        "size": item["size"],
        "encrypted": item["encrypted"],
        "source_path": item["source_path"],
    }


def _unlock_result(item: dict[str, Any], success: bool, error: str | None):
    return {
        "file_id": item["file_id"],
        "filename": item["filename"],
        "success": success,
        "error": error,
    }


def _serialize_plan(plan: MergePlan) -> dict[str, Any]:
    return {
        "sheets": [
            {
                "sheet_name": sheet.sheet_name,
                "header_row": sheet.header_row,
                "union_headers": sheet.union_headers,
                "base_file": sheet.base_file,
                "missing_files": sheet.missing_files,
                "missing_fields": sheet.missing_fields,
                "extra_fields": sheet.extra_fields,
            }
            for sheet in plan.sheets
        ],
        "warnings": plan.warnings,
    }


def _serialize_merge_summary(summary: MergeSummary) -> dict[str, Any]:
    return asdict(_to_progress_summary(summary))


@dataclass(frozen=True, slots=True)
class _ProgressSheetResult:
    """形状对齐 MergeSheetResult 的可序列化结果。

    output_files 仅为占位：共享的拆分进度接口 /api/progress 序列化时
    会访问每个 result 的 output_files，这里给空列表即可兼容。
    """

    sheet_name: str
    merged_rows: int
    source_rows: dict[str, int]
    missing_fields: dict[str, list[str]]
    extra_fields: dict[str, list[str]]
    warnings: list[str]
    skipped_duplicates: dict[str, str] = field(default_factory=dict)
    output_files: list = field(default_factory=list)


@dataclass(frozen=True, slots=True)
class _ProgressSummary:
    results: list[_ProgressSheetResult]
    output_file: str
    total_rows: int
    warnings: list[str]
    errors: list[str]


def _to_progress_summary(summary: MergeSummary) -> _ProgressSummary:
    return _ProgressSummary(
        results=[
            _ProgressSheetResult(
                sheet_name=result.sheet_name,
                merged_rows=result.merged_rows,
                source_rows=dict(result.source_rows),
                missing_fields={k: list(v) for k, v in result.missing_fields.items()},
                extra_fields={k: list(v) for k, v in result.extra_fields.items()},
                warnings=list(result.warnings),
                skipped_duplicates=dict(result.skipped_duplicates),
            )
            for result in summary.results
        ],
        output_file=str(summary.output_file),
        total_rows=summary.total_rows,
        warnings=list(summary.warnings),
        errors=list(summary.errors),
    )


def request_files() -> list:
    return [item for item in request.files.getlist("files") if item]


def request_form_text(key: str) -> str:
    value = request.form.get(key)
    return value.strip() if isinstance(value, str) else ""


def request_form_list(key: str) -> list[str]:
    return request.form.getlist(key)
