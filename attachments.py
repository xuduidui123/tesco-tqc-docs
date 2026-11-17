"""按数据源中的货号和报告号归档附件。"""
from pathlib import Path
import re
import shutil


def validate_filename(value):
    if value in {"", ".", ".."} or "/" in value or "\\" in value or any(ord(c) < 32 for c in value):
        raise ValueError(f"无效的文件或目录名称 {value!r}")


def list_files(folder):
    return sorted(p for p in Path(folder).iterdir() if p.is_file() and not p.name.startswith(('.', '~$')))


def starts_with_code(stem, code):
    stem, code = stem.casefold(), code.casefold()
    if not code or not stem.startswith(code):
        return False
    tail = stem[len(code):]
    return not tail or not tail[0].isascii() or not tail[0].isalnum()


def copy_artwork(files, sku, destination):
    matches = [p for p in files if starts_with_code(p.stem, sku.strip())]
    for path in matches:
        shutil.copy2(path, destination)
    return [] if matches else [f"未找到 Artwork（货号 {sku or '为空'}）"]


def copy_reports(files, raw_numbers, output_name, destination):
    raw_numbers = raw_numbers.strip()
    tokens = re.split(r'[\r\n,，;；、/]+', raw_numbers)
    numbers = []
    seen = set()
    for token in tokens:
        number = re.sub(r'\.pdf$', '', token.strip(), flags=re.I).strip()
        if number and number.casefold() not in seen:
            numbers.append(number)
            seen.add(number.casefold())

    issues = []
    base = re.sub(r'\.pdf$', '', output_name.strip(), flags=re.I)
    for number in numbers:
        exact = [p for p in files if p.stem.casefold() == number.casefold()]
        matches = exact or [p for p in files if starts_with_code(p.stem, number)]
        if not matches:
            issues.append(f"缺少检测报告 {number}")
            continue
        if len(matches) > 1:
            issues.append(f"报告 {number} 有多个匹配，请填写具体文件名：" + "、".join(p.name for p in matches))
            continue
        if not base:
            name = matches[0].stem
        elif len(numbers) == 1 and len([t for t in tokens if t.strip()]) == 1:
            name = base
        elif base.endswith(raw_numbers):
            # 数据源自带的命名公式以整格报告号结尾。
            name = base.removesuffix(raw_numbers) + number
        else:
            name = f"{base}-{number}"
        validate_filename(name)
        shutil.copy2(matches[0], Path(destination) / (name + ".pdf"))
    return issues
