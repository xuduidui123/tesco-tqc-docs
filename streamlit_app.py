import io
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st

from main import (EXCEL_FILE, IMAGE_DIR, OUTPUT_BASE_DIR, PUBLIC_CERTS_DIR,
                  TEST_REPORTS_DIR, generate, validate_filename)


def build_download(workbook, groups):
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / EXCEL_FILE).write_bytes(workbook.getbuffer())
        for dirname, files in groups:
            folder = root / dirname
            folder.mkdir()
            for upload in files:
                validate_filename(upload.name)
                target = folder / upload.name
                if target.exists():
                    raise ValueError(f"重复上传了同名文件 {upload.name}")
                target.write_bytes(upload.getbuffer())
        issues = generate(root, cloud=True)
        output = root / OUTPUT_BASE_DIR
        files = sorted(path for path in output.rglob("*") if path.is_file())
        archive = io.BytesIO()
        with ZipFile(archive, "w", ZIP_DEFLATED) as zip_file:
            for path in files:
                zip_file.write(path, path.relative_to(output))
        return archive.getvalue(), len(files), issues


def clear_result():
    st.session_state.pop("result", None)


def app():
    st.set_page_config(page_title="TQC 文件生成")
    st.title("TQC 文件生成")
    workbook = st.file_uploader("TQC 数据源 Excel", type=["xlsx"], on_change=clear_result)
    images = st.file_uploader("产品图片", type=["png", "jpg", "jpeg"], accept_multiple_files=True,
                              help="文件名须与 Excel 的“产品图片”列一致。", on_change=clear_result)
    reports = st.file_uploader("检测报告", type=["pdf"], accept_multiple_files=True,
                               help="文件名为检测报告号加 .pdf。", on_change=clear_result)
    certs = st.file_uploader("FSC 证书", type=["pdf"], accept_multiple_files=True,
                             help="按需上传，文件名须与“FSC证书类型”列一致。", on_change=clear_result)
    if st.button("生成 TQC 文件", type="primary", disabled=workbook is None):
        clear_result()
        try:
            with st.spinner("正在填充模板并转换 PDF…"):
                st.session_state.result = build_download(workbook, [
                    (IMAGE_DIR, images), (TEST_REPORTS_DIR, reports), (PUBLIC_CERTS_DIR, certs)])
        except Exception as exc:
            st.error(f"生成失败。{exc}")
    if "result" in st.session_state:
        data, count, issues = st.session_state.result
        if issues:
            st.warning("部分附件未归档，请核对后再交付。")
            st.text("\n".join(issues))
        st.write(f"本批次共生成或归档 {count} 个文件。")
        st.download_button("下载 ZIP", data, "TQC-documents.zip", "application/zip")


if __name__ == "__main__":
    app()
