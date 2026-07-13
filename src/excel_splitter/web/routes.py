from __future__ import annotations

from dataclasses import asdict
from datetime import date, datetime, time
from pathlib import Path
from threading import Thread
from typing import Any
from uuid import uuid4

from flask import Blueprint, current_app, jsonify, render_template, request, send_file, url_for

from ..engine import SplitEngine
from ..models import SheetConfig, SplitJob
from ..planning import build_split_plan
from ..preview import list_sheet_names, preview_sheet
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
    try:
        sheets = list_sheet_names(stored_path)
    except Exception:
        stored_path.unlink(missing_ok=True)
        raise ValueError("文件不是可读取的 .xlsx 工作簿")

    _jobs()[job_id] = {
        "input_file": stored_path,
        "original_name": Path(original_filename).stem,
        "results": [],
    }
    return jsonify(
        job_id=job_id,
        filename=original_filename,
        sheets=sheets,
        default_output_dir=str(current_app.config["DEFAULT_OUTPUT_DIR"]),
    )


@api.post("/api/preview")
def preview():
    payload = _json_payload()
    job = _get_job(payload.get("job_id"))
    result = preview_sheet(
        job["input_file"],
        _required_text(payload, "sheet_name"),
        start_row=int(payload.get("start_row", 1)),
        max_rows=int(payload.get("max_rows", 100)),
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
    plan = build_split_plan(job["input_file"], configs)
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
    split_job = SplitJob(
        input_file=job_record["input_file"],
        output_dir=output_dir,
        sheet_configs=_sheet_configs(payload.get("sheet_configs")),
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
            args=(job_record, split_job),
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

    summary = SplitEngine().execute(split_job)
    _store_results(job_record, summary)
    return jsonify(_serialize_summary(job_id, summary))


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
        payload["result"] = _serialize_summary(job_id, execution["summary"])
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


def _run_background_job(job_record: dict[str, Any], split_job: SplitJob) -> None:
    execution = job_record["execution"]

    def update_progress(percent: int, message: str) -> None:
        execution.update(
            status="running",
            progress=percent,
            message=message,
        )

    try:
        summary = SplitEngine().execute(
            split_job, progress_callback=update_progress
        )
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


def _store_results(job_record: dict[str, Any], summary) -> None:
    job_record["results"] = [
        artifact.output_file
        for result in summary.results
        for artifact in result.output_files
    ]


def _serialize_summary(job_id: str, summary) -> dict[str, Any]:
    result_payload = asdict(summary)
    download_index = 0
    for result in result_payload["results"]:
        for artifact in result["output_files"]:
            artifact["output_file"] = str(artifact["output_file"])
            artifact["download_url"] = url_for(
                "excel_splitter.download",
                job_id=job_id,
                result_index=download_index,
            )
            download_index += 1
    return result_payload


@api.get("/api/result/<job_id>")
def result(job_id: str):
    job = _get_job(job_id)
    return jsonify(
        files=[
            {
                "output_file": str(path),
                "download_url": url_for(
                    "excel_splitter.download", job_id=job_id, result_index=index
                ),
            }
            for index, path in enumerate(job["results"])
        ]
    )


@api.get("/api/download/<job_id>/<int:result_index>")
def download(job_id: str, result_index: int):
    job = _get_job(job_id)
    try:
        output_file = job["results"][result_index]
    except IndexError as exc:
        raise ValueError("找不到导出文件") from exc
    if not output_file.is_file():
        raise ValueError("导出文件已不存在")
    return send_file(
        output_file,
        as_attachment=True,
        download_name=output_file.name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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
