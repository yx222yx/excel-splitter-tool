from pathlib import Path
import tomllib


ROOT = Path(__file__).resolve().parents[1]


def test_project_declares_desktop_and_build_dependencies():
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))

    dependencies = set(project["project"]["dependencies"])
    build_dependencies = set(project["project"]["optional-dependencies"]["build"])

    assert "pywebview==6.2.1" in dependencies
    assert "waitress==3.0.2" in dependencies
    assert "pyinstaller==6.21.0" in build_dependencies
    assert (
        project["project"]["scripts"]["excel-splitter-desktop"]
        == "excel_splitter.desktop:main"
    )


def test_portable_build_assets_exist_and_target_onedir_windowed_app():
    spec = (ROOT / "packaging" / "excel_splitter.spec").read_text(encoding="utf-8")
    build_script = (ROOT / "scripts" / "build_portable.ps1").read_text(
        encoding="utf-8"
    )
    user_guide = ROOT / "packaging" / "使用说明.txt"
    version_file = ROOT / "packaging" / "version.txt"

    assert "collect_data_files('excel_splitter.web')" in spec
    assert "webview.platforms.edgechromium" in spec
    assert "webview.platforms.winforms" in spec
    assert "collect_submodules('webview')" not in spec
    assert "name='Excel拆分工具'" in spec
    assert "console=False" in spec
    assert "COLLECT(" in spec
    assert "PyInstaller" in build_script
    assert "excel_splitter.spec" in build_script
    assert "--distpath" in build_script
    assert "--workpath" in build_script
    assert '"Library\\bin"' in build_script
    assert "Get-ChildItem" in build_script
    assert "Remove-Item -LiteralPath $DistRoot -Recurse -Force" in build_script
    assert "3.11|64" in build_script
    assert "Compress-Archive -Path $DistDir.FullName" in build_script
    assert user_guide.exists()
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert version_file.read_text(encoding="ascii").strip() == project["project"]["version"]


def test_readme_names_the_generated_portable_zip():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")

    assert "ExcelSplitter-portable.zip" in readme
