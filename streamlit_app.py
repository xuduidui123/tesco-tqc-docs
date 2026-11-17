import io
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st

from main import (ARTWORK_DIR, EXCEL_FILE, IMAGE_DIR, OUTPUT_BASE_DIR, PUBLIC_CERTS_DIR,
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
    st.set_page_config(page_title="TQC 文件生成（优化版）")
    st.title("TQC 文件生成（优化版）")
    workbook = st.file_uploader("TQC 数据源 Excel", type=["xlsx"], on_change=clear_result)
    images = st.file_uploader("产品图片", type=["png", "jpg", "jpeg"], accept_multiple_files=True,
                              help="文件名须与 Excel 的“产品图片”列一致。", on_change=clear_result)
    artwork = st.file_uploader("Artwork 文件", accept_multiple_files=True, on_change=clear_result,
                              help="文件名以完整货号开头。同一货号的多份 Artwork 会一起归档。")
    reports = st.file_uploader("检测报告", type=["pdf"], accept_multiple_files=True,
                               help="支持一款多份报告和多款共用报告。报告号可用换行、逗号或斜线分隔；同号多版本请填完整文件名。", on_change=clear_result)
    certs = st.file_uploader("FSC 证书", type=["pdf"], accept_multiple_files=True,
                             help="按需上传，文件名须与“FSC证书类型”列一致。", on_change=clear_result)
    if st.button("生成 TQC 文件", type="primary", disabled=workbook is None):
        clear_result()
        try:
            with st.spinner("正在填充模板并转换 PDF…"):
                st.session_state.result = build_download(workbook, [
                    (IMAGE_DIR, images), (ARTWORK_DIR, artwork), (TEST_REPORTS_DIR, reports), (PUBLIC_CERTS_DIR, certs)])
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
