from __future__ import annotations

import os
import subprocess
from contextlib import nullcontext
from dataclasses import asdict
from datetime import date, datetime, time
from pathlib import Path
from threading import Lock, Thread
from typing import Any
from uuid import uuid4

from flask import Blueprint, current_app, jsonify, render_template, request, url_for

from ..encryption import decrypt_file, encrypt_file, is_encrypted
from ..engine import SplitEngine
from ..excel_io import load_workbook_with_warnings
from ..models import SheetConfig, SplitJob
from ..planning import SplitPlan, build_split_plan
from ..preview import preview_sheet
from .dialogs import choose_directory


api = Blueprint("excel_splitter", __name__)


@api.get("/")
def index():
    return render_template("index.html")


@api.post("/api/load")
def load_file():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        raise ValueError("请选择 Excel 文件")
    original_filename = Path(upload.filename).name
    if Path(original_filename).suffix.lower() != ".xlsx":
        raise ValueError("第一版仅支持 .xlsx 文件")

    job_id = uuid4().hex
    stored_path = current_app.config["UPLOAD_DIR"] / f"{job_id}.xlsx"
    upload.save(stored_path)

    source_dir = request.form.get("source_dir")

    if is_encrypted(stored_path):
        _jobs()[job_id] = {
            "input_file": stored_path,
            "original_name": Path(original_filename).stem,
            "encrypted": True,
            "source_dir": source_dir or "",
        }
        return jsonify(needs_password=True, job_id=job_id, filename=original_filename)

    try:
        workbook, _ = load_workbook_with_warnings(stored_path, data_only=True)
        sheets = list(workbook.sheetnames)
    except Exception:
        stored_path.unlink(missing_ok=True)
        raise ValueError("文件不是可读取的 .xlsx 工作簿")

    if source_dir:
        default_output_dir = str(
            Path(source_dir) / f"{Path(original_filename).stem}_拆分结果"
        )
    else:
        default_output_dir = str(current_app.config["DEFAULT_OUTPUT_DIR"])

    _jobs()[job_id] = {
        "input_file": stored_path,
        "original_name": Path(original_filename).stem,
        "results": [],
        "_workbook": workbook,
        "_workbook_lock": Lock(),
        "encrypted": False,
    }
    return jsonify(
        job_id=job_id,
        filename=original_filename,
        sheets=sheets,
        default_output_dir=default_output_dir,
    )


@api.post("/api/load-with-password")
def load_with_password():
    payload = _json_payload()
    job_id = _required_text(payload, "job_id")
    password = _required_text(payload, "password")
    job = _get_job(job_id)
    if not job.get("encrypted"):
        raise ValueError("文件未加密，无需密码")

    stored_path = Path(job["input_file"])
    try:
        decrypted = decrypt_file(stored_path, password)
    except Exception:
        raise ValueError("密码错误，无法解密文件")

    decrypted_path = stored_path.with_suffix(".decrypted.xlsx")
    with open(decrypted_path, "wb") as f:
        f.write(decrypted.read())

    try:
        workbook, _ = load_workbook_with_warnings(decrypted_path, data_only=True)
        sheets = list(workbook.sheetnames)
    except Exception:
        decrypted_path.unlink(missing_ok=True)
        raise ValueError("解密后文件无法读取")

    original_filename = Path(job["original_name"] + ".xlsx").name
    job.update({
        "input_file": decrypted_path,
        "results": [],
        "_workbook": workbook,
        "_workbook_lock": Lock(),
        "encrypted": False,
    })

    source_dir = job.get("source_dir", "")
    if source_dir:
        default_output_dir = str(
            Path(source_dir) / f"{Path(original_filename).stem}_拆分结果"
        )
    else:
        default_output_dir = str(current_app.config["DEFAULT_OUTPUT_DIR"])

    return jsonify(
        job_id=job_id,
        filename=original_filename,
        sheets=sheets,
        default_output_dir=default_output_dir,
    )


@api.post("/api/preview")
def preview():
    payload = _json_payload()
    job = _get_job(payload.get("job_id"))
    wb, lock = job.get("_workbook"), job.get("_workbook_lock")
    with lock or nullcontext():
        result = preview_sheet(
            job["input_file"],
            _required_text(payload, "sheet_name"),
            start_row=int(payload.get("start_row", 1)),
            max_rows=int(payload.get("max_rows", 100)),
            workbook=wb,
        )
    return jsonify(
        sheet_name=result.sheet_name,
        rows=[[_json_value(value) for value in row] for row in result.rows],
        suggested_header_row=result.suggested_header_row,
        start_row=result.start_row,
        end_row=result.end_row,
        total_rows=result.total_rows,
        has_more=result.has_more,
        warnings=list(result.warnings),
    )


@api.post("/api/split-values")
def split_values():
    payload = _json_payload()
    job = _get_job(payload.get("job_id"))
    configs = _sheet_configs(payload.get("sheet_configs"))
    wb, lock = job.get("_workbook"), job.get("_workbook_lock")
    with lock or nullcontext():
        plan = build_split_plan(job["input_file"], configs, workbook=wb)
    job["_split_plan"] = plan
    job["_split_plan_configs"] = configs
    return jsonify(
        values=plan.values, empty_rows=plan.empty_rows, warnings=plan.warnings
    )


@api.post("/api/execute")
def execute():
    payload = _json_payload()
    job_record = _get_job(payload.get("job_id"))
    job_id = payload["job_id"]
    output_dir = Path(_required_text(payload, "output_dir"))
    if not output_dir.is_absolute():
        raise ValueError("输出目录必须是本机绝对路径")
    raw_output_types = payload.get("output_types")
    if raw_output_types is None:
        output_types = ("formula", "values")
    elif isinstance(raw_output_types, list):
        output_types = tuple(raw_output_types)
    else:
        raise ValueError("输出版本必须是列表")
    configs = _sheet_configs(payload.get("sheet_configs"))
    split_job = SplitJob(
        input_file=job_record["input_file"],
        output_dir=output_dir,
        sheet_configs=configs,
        split_mode=payload.get("split_mode", "all"),
        selected_split_values=tuple(payload.get("selected_split_values") or ()),
        filename_template=payload.get(
            "filename_template", "{original_name}_{split_value}_{output_type}"
        ),
        output_types=output_types,
        overwrite=bool(payload.get("overwrite", False)),
        original_name=job_record["original_name"],
    )
    split_job.validate()
    output_password = payload.get("output_password", "") or ""
    cached_plan = (
        job_record.get("_split_plan")
        if job_record.get("_split_plan_configs") == configs
        else None
    )

    if bool(payload.get("background", False)):
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
            target=_run_background_job,
            args=(job_record, split_job, output_password, cached_plan),
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

    summary = SplitEngine().execute(split_job, plan=cached_plan)
    if output_password:
        _encrypt_summary_outputs(summary, output_password)
    _store_results(job_record, summary)
    return jsonify(_serialize_summary(summary))


@api.get("/api/progress/<job_id>")
def execution_progress(job_id: str):
    job_record = _get_job(job_id)
    execution = job_record.get("execution")
    if not execution:
        raise ValueError("任务尚未开始执行")
    payload = {
        "status": execution["status"],
        "progress": execution["progress"],
        "message": execution["message"],
    }
    if execution["status"] == "complete":
        payload["result"] = _serialize_summary(execution["summary"])
    elif execution["status"] == "failed":
        payload["error"] = execution["error"]
    return jsonify(payload)


@api.post("/api/select-output-dir")
def select_output_dir():
    payload = _json_payload()
    current_path = payload.get("current_path")
    initial_dir = Path(current_path) if isinstance(current_path, str) else None
    if initial_dir is None or not initial_dir.is_dir():
        initial_dir = current_app.config["DEFAULT_OUTPUT_DIR"]
    selected = choose_directory(initial_dir)
    if selected is None:
        return jsonify(selected=False, path=None)
    return jsonify(selected=True, path=str(selected))


def _run_background_job(
    job_record: dict[str, Any],
    split_job: SplitJob,
    output_password: str = "",
    plan: SplitPlan | None = None,
) -> None:
    execution = job_record["execution"]

    def update_progress(percent: int, message: str) -> None:
        execution.update(
            status="running",
            progress=percent,
            message=message,
        )

    try:
        summary = SplitEngine().execute(
            split_job, progress_callback=update_progress, plan=plan
        )
        if output_password:
            _encrypt_summary_outputs(summary, output_password)
        _store_results(job_record, summary)
        execution.update(
            status="complete",
            progress=100,
            message="拆分完成",
            summary=summary,
        )
    except Exception as exc:
        execution.update(
            status="failed",
            message="拆分失败",
            error=str(exc),
        )


def _encrypt_summary_outputs(summary, password: str) -> None:
    """加密输出结果中的所有文件。"""
    for result in summary.results:
        for artifact in result.output_files:
            encrypt_file(artifact.output_file, password)


def _store_results(job_record: dict[str, Any], summary) -> None:
    job_record["results"] = [
        artifact.output_file
        for result in summary.results
        for artifact in result.output_files
    ]


def _serialize_summary(summary) -> dict[str, Any]:
    result_payload = asdict(summary)
    for result in result_payload["results"]:
        for artifact in result["output_files"]:
            artifact["output_file"] = str(artifact["output_file"])
    return result_payload


@api.post("/api/action/open-file")
def open_file():
    payload = _json_payload()
    path = Path(_required_text(payload, "path"))
    if not path.is_file():
        raise ValueError("文件不存在")
    os.startfile(str(path))
    return jsonify(ok=True)


@api.post("/api/action/open-folder")
def open_folder():
    payload = _json_payload()
    path = Path(_required_text(payload, "path"))
    if not path.is_file():
        raise ValueError("文件不存在")
    subprocess.Popen(["explorer", "/select,", str(path)])
    return jsonify(ok=True)


def _jobs() -> dict[str, dict[str, Any]]:
    return current_app.extensions["excel_splitter_jobs"]


def _get_job(job_id: Any) -> dict[str, Any]:
    if not isinstance(job_id, str) or job_id not in _jobs():
        raise ValueError("任务不存在或已过期，请重新选择文件")
    return _jobs()[job_id]


def _json_payload() -> dict[str, Any]:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        raise ValueError("请求内容必须是 JSON 对象")
    return payload


def _required_text(payload: dict[str, Any], key: str) -> str:
    value = payload.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"缺少参数：{key}")
    return value.strip()


def _sheet_configs(raw_configs: Any) -> tuple[SheetConfig, ...]:
    if not isinstance(raw_configs, list) or not raw_configs:
        raise ValueError("至少配置一个 sheet")
    configs: list[SheetConfig] = []
    for raw in raw_configs:
        if not isinstance(raw, dict):
            raise ValueError("sheet 配置格式无效")
        try:
            config = SheetConfig(
                sheet_name=str(raw["sheet_name"]),
                header_row=int(raw["header_row"]),
                split_column_idx=_optional_int(raw.get("split_column_idx")),
                split_column_label=str(raw.get("split_column_label", "")),
                mode=str(raw.get("mode", "direct")),
                key_column_idx=_optional_int(raw.get("key_column_idx")),
                key_column_label=str(raw.get("key_column_label", "")),
            )
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError("sheet 配置缺少有效的表头行或拆分列") from exc
        config.validate()
        configs.append(config)
    return tuple(configs)


def _optional_int(value: Any) -> int | None:
    if value is None or value == "":
        return None
    return int(value)


def _json_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (date, datetime, time)):
        return value.isoformat()
    return str(value)
