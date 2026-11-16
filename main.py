import os
import shutil
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import subprocess
import tempfile
from pathlib import Path
from tqdm import tqdm
import openpyxl 
from PIL import Image as PILImage 
from openpyxl.drawing.image import Image as ExcelImage 

# ================= 1. 基础路径配置 =================
# 从终端或编辑器启动时，也按 main.py 所在目录寻找模板和数据。
PROJECT = Path(__file__).resolve().parent

EXCEL_FILE = "4. TQC文件数据源.xlsx"
IMAGE_DIR = "1. Image"
TEST_REPORTS_DIR = "2. Test Report"  
PUBLIC_CERTS_DIR = "3. Public_Certs"   
OUTPUT_BASE_DIR = "5. output"

SA_FORM_TEMPLATE = "模板_SA Form.xlsx" 

MAC_HOME = os.path.expanduser("~")
OFFICE_SAFE_ZONE = os.path.join(MAC_HOME, "Library/Group Containers/UBF8T346G9.Office/Tesco_Temp")

TEMPLATES_CONFIG = [
    {"template_name": "模板_Tesco_DO_CE.docx", "name_col": "Tesco DOC EU名字", "need_pdf": False},
    {"template_name": "模板_Tesco_DO_UK.docx", "name_col": "Tesco DOC UK名字", "need_pdf": False},
    {"template_name": "模板_Supplier_DO_CE.docx", "name_col": "Supplier DOC EU名字", "need_pdf": True},
    {"template_name": "模板_Supplier_DO_UK.docx", "name_col": "Supplier DOC UK名字", "need_pdf": True},
    {"template_name": "模板_Reach DOC.docx", "name_col": None, "fixed_name": "Reach DOC", "need_pdf": True},
    {"template_name": "模板_DOO.docx", "name_col": "DOO名字", "need_pdf": False} 
]

# ================= 2. SA Form 填充函数 (包含精确的英寸尺寸限制) =================
def fill_sa_form(template_path, output_path, data_dict, img_path=None):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    safe_data_dict = {str(k): str(v) for k, v in data_dict.items()}
    
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                
                # 专门拦截图片占位符
                if "{{产品图片}}" in cell.value:
                    cell.value = cell.value.replace("{{产品图片}}", "") 
                    if img_path and os.path.exists(img_path):
                        try:
                            img = ExcelImage(img_path)
                            
                            # 【核心升级】：精确英寸到像素的换算限制
                            # 1 inch ≈ 96 pixels
                            MAX_W_PX = int(3.6 * 96) # 约 345 像素
                            MAX_H_PX = int(3.0 * 96) # 约 288 像素
                            
                            with PILImage.open(img_path) as pil_img:
                                orig_w, orig_h = pil_img.size
                                ratio = orig_w / orig_h
                                limit_ratio = MAX_W_PX / MAX_H_PX
                                
                                # 智能等比缩放，绝不超出给定的宽高限制
                                if ratio > limit_ratio:
                                    img.width = MAX_W_PX
                                    img.height = int(MAX_W_PX / ratio)
                                else:
                                    img.height = MAX_H_PX
                                    img.width = int(MAX_H_PX * ratio)
                                    
                            ws.add_image(img, cell.coordinate) 
                        except Exception as e:
                            pass
                    continue 
                
                for key, value in safe_data_dict.items():
                    if key == "产品图片": continue 
                    placeholder = f"{{{{{key}}}}}"
                    if cell.value and isinstance(cell.value, str) and placeholder in cell.value:
                        cell.value = cell.value.replace(placeholder, value)
    wb.save(output_path)

def validate_filename(value):
    if value in {"", ".", ".."} or "/" in value or "\\" in value or any(ord(c) < 32 for c in value):
        raise ValueError(f"无效的文件或目录名称 {value!r}")


def generate_row(row, root, scratch, profile, cloud, warnings):
    raw_context = {str(k): v for k, v in row.to_dict().items()}
    if cloud:
        for field in ("序号", "中文品名", "文件夹名字", "产品图片", "检测报告号",
                      "检测报告名字", "英文品名", *(cfg["name_col"] for cfg in TEMPLATES_CONFIG if cfg["name_col"])):
            value = raw_context.get(field, "").strip()
            if value:
                validate_filename(value)
    render_context = {k: v.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;') for k, v in raw_context.items()}
    
    # --- 目录构建 ---
    seq_num = raw_context.get('序号', '').strip()
    parent_dir = os.path.join(root / OUTPUT_BASE_DIR, f"{seq_num}. {raw_context.get('中文品名')}")
    child_dir = os.path.join(parent_dir, str(raw_context.get('文件夹名字')).strip())
    if not os.path.exists(child_dir): os.makedirs(child_dir)
    


    # --- 功能①：复制产品原图 ---
    img_name = str(raw_context.get('产品图片', '')).strip()
    src_img = os.path.join(root / IMAGE_DIR, img_name)
    if os.path.exists(src_img) and img_name:
        ext = os.path.splitext(img_name)[1]
        shutil.copy(src_img, os.path.join(parent_dir, f"1. Product Image{ext}"))

    elif img_name:
        warnings.append(f"缺少产品图 {img_name}")

    # --- 功能②：抓取检测报告 ---
    report_no = str(raw_context.get('检测报告号', '')).strip()
    src_report = os.path.join(root / TEST_REPORTS_DIR, f"{report_no}.pdf")
    final_report_name_base = str(raw_context.get('检测报告名字', '')).strip()
    if os.path.exists(src_report) and final_report_name_base:
        shutil.copy(src_report, os.path.join(child_dir, f"{final_report_name_base}.pdf"))

    elif report_no or final_report_name_base:
        warnings.append(f"检测报告未归档，请核对报告号及输出名称（{report_no or '未填报告号'}）")

    # --- 功能③：智能分发公共证书 ---
    fsc_files_str = str(raw_context.get('FSC证书类型', '')).strip()
    if fsc_files_str:
        fsc_files = [f.strip() for f in fsc_files_str.replace('\n', ',').split(',') if f.strip()]
        for i, f_name in enumerate(fsc_files):
            if cloud:
                validate_filename(f_name)
            src_fsc = os.path.join(root / PUBLIC_CERTS_DIR, f_name)
            if not os.path.isfile(src_fsc) and not f_name.lower().endswith(".pdf"):
                src_fsc += ".pdf"
            if os.path.isfile(src_fsc):
                idx = f"5.{i+1}" if len(fsc_files) > 1 else "5"
                shutil.copy(src_fsc, os.path.join(parent_dir, f"{idx}. FSC Certificate.pdf"))

            else:
                warnings.append(f"缺少证书 {f_name}")

    # --- 解析 Test Program ---
    test_program_raw = str(raw_context.get('Test Program', '')).strip()
    test_lines = [line.strip() for line in test_program_raw.split('\n') if line.strip()]

    for i in range(1, 9): raw_context[f'Test_{i}'] = ""
    for i, line in enumerate(test_lines):
        if i < 7: raw_context[f'Test_{i+1}'] = line
        else:
            raw_context['Test_8'] = f"{raw_context['Test_8']}\n{line}" if raw_context['Test_8'] else line

    # --- 功能④：生成 SA Form Excel ---
    function_val = str(raw_context.get('Function:', '')).strip()
    if function_val:
        en_name = str(raw_context.get('英文品名', '')).strip()
        if en_name:
            sa_name = f"SA FORM-{en_name}.xlsx"
            img_path_for_excel = src_img if (os.path.exists(src_img) and img_name) else None
            fill_sa_form(PROJECT / SA_FORM_TEMPLATE, os.path.join(parent_dir, sa_name), raw_context, img_path_for_excel)

    # --- 生成 Word/PDF ---
    for config in TEMPLATES_CONFIG:
        template_path = PROJECT / config["template_name"]
        if not template_path.is_file():
            raise FileNotFoundError(f"缺少模板 {template_path.name}")
            
        target_file_name_base = config.get("fixed_name") or str(raw_context.get(config["name_col"], "")).strip()
        if not target_file_name_base: continue
            
        doc = DocxTemplate(template_path)
        
        if os.path.exists(src_img) and img_name:
            try:
                with PILImage.open(src_img) as img:
                    w, h = img.size
                    ratio = w / h
                    fit_args = {'width': Mm(60)} if ratio < (60/80) else {'width': Mm(60), 'height': Mm(60/ratio)}
                render_context['产品图片'] = InlineImage(doc, src_img, **fit_args)
            except Exception:
                render_context['产品图片'] = InlineImage(doc, src_img, width=Mm(60))
        else:
            render_context['产品图片'] = "（无图片）"
            
        current_save_dir = parent_dir if config.get("fixed_name") == "Reach DOC" else child_dir
            
        if config["need_pdf"]:
            with tempfile.TemporaryDirectory(dir=scratch) as tmp:
                word = Path(tmp) / "document.docx"
                pdf = word.with_suffix(".pdf")
                doc.render(render_context)
                doc.save(word)
                if cloud:
                    subprocess.run([
                        "soffice", f"-env:UserInstallation={profile.as_uri()}", "--headless",
                        "--convert-to", "pdf", "--outdir", tmp, str(word),
                    ], check=True, capture_output=True, timeout=120)
                else:
                    from docx2pdf import convert
                    convert(str(word), str(pdf))
                if not pdf.is_file() or not pdf.stat().st_size:
                    raise RuntimeError(f"{target_file_name_base}.pdf 转换失败")
                shutil.move(pdf, Path(current_save_dir) / f"{target_file_name_base}.pdf")
        else:
            target_word = os.path.join(current_save_dir, f"{target_file_name_base}.docx")
            doc.render(render_context)
            doc.save(target_word)



def generate(root, *, cloud=False):
    root = Path(root).resolve()
    scratch = root if cloud else Path(OFFICE_SAFE_ZONE)
    scratch.mkdir(parents=True, exist_ok=True)
    for directory in (OUTPUT_BASE_DIR, TEST_REPORTS_DIR, PUBLIC_CERTS_DIR):
        (root / directory).mkdir(exist_ok=True)
    df = pd.read_excel(root / EXCEL_FILE, dtype=str, engine="openpyxl", keep_default_na=False)
    if "文件夹名字" not in df:
        raise ValueError("Excel 缺少“文件夹名字”列，请使用 TQC 数据源表。")
    active = df[df["文件夹名字"].str.strip() != ""]
    if active.empty:
        raise ValueError("Excel 中没有填写“文件夹名字”的产品。")
    issues = []
    for index, row in tqdm(active.iterrows(), total=len(active), disable=cloud):
        warnings = []
        try:
            generate_row(row, root, scratch, root / "lo-profile", cloud, warnings)
        except Exception as exc:
            raise RuntimeError(f"Excel 第 {index + 2} 行（{row.get('中文品名', '')}）生成失败：{exc}") from exc
        issues.extend(f"第 {index + 2} 行：{msg}" for msg in warnings)
    return issues


if __name__ == "__main__":
    for issue in generate(PROJECT):
        print(issue)
    print("已完成")
