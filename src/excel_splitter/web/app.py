from __future__ import annotations

import atexit
from pathlib import Path
import shutil
import tempfile

from flask import Flask, jsonify, request

from .merge_routes import merge_api
from .routes import api


def create_app(config: dict | None = None) -> Flask:
    app = Flask(__name__)
    project_root = Path(__file__).resolve().parents[3]
    generated_upload_dir = Path(tempfile.mkdtemp(prefix="excel-splitter-"))
    app.config.from_mapping(
        MAX_CONTENT_LENGTH=100 * 1024 * 1024,
        UPLOAD_DIR=generated_upload_dir,
        DEFAULT_OUTPUT_DIR=Path.home() / "Desktop",
        JSON_AS_ASCII=False,
    )
    if config:
        app.config.update(config)

    app.config["UPLOAD_DIR"] = Path(app.config["UPLOAD_DIR"])
    app.config["DEFAULT_OUTPUT_DIR"] = Path(app.config["DEFAULT_OUTPUT_DIR"])
    app.config["UPLOAD_DIR"].mkdir(parents=True, exist_ok=True)
    app.config["DEFAULT_OUTPUT_DIR"].mkdir(parents=True, exist_ok=True)
    app.extensions["excel_splitter_jobs"] = {}
    app.register_blueprint(api)
    app.register_blueprint(merge_api)

    if app.config["UPLOAD_DIR"] == generated_upload_dir:
        atexit.register(shutil.rmtree, generated_upload_dir, True)
    else:
        shutil.rmtree(generated_upload_dir, ignore_errors=True)

    @app.errorhandler(ValueError)
    def handle_value_error(error):
        return jsonify(error=str(error)), 400

    @app.errorhandler(413)
    def handle_too_large(_error):
        return jsonify(error="文件超过 100 MB 限制"), 413

    @app.errorhandler(Exception)
    def handle_unexpected_error(error):
        if app.config["TESTING"]:
            raise error
        if request.path.startswith("/api/"):
            app.logger.exception("API request failed")
            return jsonify(error="处理失败，请查看服务日志"), 500
        raise error

    return app


def main() -> None:
    create_app().run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()

