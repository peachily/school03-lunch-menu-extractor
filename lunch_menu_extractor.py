import os
import re
import zipfile
import sys
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import filedialog, messagebox


NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
}

WEEKDAY_MAP = {
    "월": "월요일",
    "화": "화요일",
    "수": "수요일",
    "목": "목요일",
    "금": "금요일",
}


def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="급식표 HWPX 파일 선택",
        filetypes=[("HWPX files", "*.hwpx")],
    )
    root.destroy()
    return file_path


def get_month_from_filename(file_path):
    base = os.path.basename(file_path)
    m = re.search(r"(\d{1,2})월", base)
    if not m:
        raise RuntimeError("파일명에서 월 정보를 찾지 못했습니다.")
    return int(m.group(1))


def build_date_re(month):
    return re.compile(rf"{month}/(\d{{1,2}})\(([월화수목금])\)")


def get_output_path(file_path, month):
    return os.path.join(os.path.dirname(file_path), f"{month}월 급식.txt")


def load_section_xml(hwpx_path):
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        xml_bytes = zf.read("Contents/section0.xml")
    return ET.fromstring(xml_bytes)


def get_tables(root):
    return root.findall(".//hp:tbl", NS)


def get_cell_addr(tc):
    addr = tc.find("./hp:cellAddr", NS)
    if addr is None:
        return None, None
    return int(addr.attrib.get("rowAddr", -1)), int(addr.attrib.get("colAddr", -1))


def get_cell_text(tc):
    parts = []
    for t in tc.findall(".//hp:t", NS):
        text = "".join(t.itertext()).strip()
        if text:
            parts.append(text)
    return " ".join(parts).strip()


def build_cell_map(table):
    cell_map = {}
    for tc in table.findall(".//hp:tc", NS):
        row, col = get_cell_addr(tc)
        if row is None or col is None or row < 0 or col < 0:
            continue

        text = get_cell_text(tc)
        if not text:
            continue

        key = (row, col)
        if key not in cell_map:
            cell_map[key] = []
        cell_map[key].append(text)

    return cell_map


def pick_date_text(candidates, date_re):
    for text in candidates:
        if date_re.search(text):
            return text
    return ""


def clean_menu_text(text):
    text = text.replace("\xa0", " ")
    text = text.replace("\n", " ")
    text = text.replace("\r", " ")
    text = text.replace("/", " ")
    text = re.sub(r"\([^)]*\)", "", text)
    text = text.replace("우유", "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def pick_menu_text(candidates, date_re):
    for text in candidates:
        if date_re.search(text):
            continue
        cleaned = clean_menu_text(text)
        if cleaned:
            return cleaned
    return ""


def extract_meals_from_table(table, date_re):
    cell_map = build_cell_map(table)
    if not cell_map:
        return []

    rows = sorted({row for row, _ in cell_map.keys()})
    results = []

    for row in rows:
        cols = sorted(col for (r, col) in cell_map.keys() if r == row)

        for col in cols:
            header_candidates = cell_map.get((row, col), [])
            header_text = pick_date_text(header_candidates, date_re)
            if not header_text:
                continue

            match = date_re.search(header_text)
            if not match:
                continue

            day = int(match.group(1))
            weekday_short = match.group(2)
            weekday = WEEKDAY_MAP[weekday_short]

            menu_candidates = cell_map.get((row + 1, col), [])
            menu_text = pick_menu_text(menu_candidates, date_re)
            if not menu_text:
                continue

            results.append((day, f"{day}일 {weekday} {menu_text}"))

    unique = {}
    for day, line in results:
        unique[day] = line

    return [unique[day] for day in sorted(unique.keys())]


def get_meal_table(root, date_re):
    tables = get_tables(root)

    matched_tables = []
    for table in tables:
        cell_map = build_cell_map(table)
        hit_count = 0

        for candidates in cell_map.values():
            if pick_date_text(candidates, date_re):
                hit_count += 1

        if hit_count > 0:
            matched_tables.append((hit_count, table))

    if not matched_tables:
        raise RuntimeError("날짜가 들어 있는 표를 찾지 못했습니다.")

    matched_tables.sort(key=lambda x: x[0], reverse=True)
    return matched_tables[0][1]


def save_result(lines, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def main():
    file_path = select_file()
    if not file_path:
        return

    if not file_path.lower().endswith(".hwpx"):
        messagebox.showerror("오류", "HWPX 파일만 선택해 주세요.")
        return

    try:
        month = get_month_from_filename(file_path)
        date_re = build_date_re(month)

        root = load_section_xml(file_path)
        table = get_meal_table(root, date_re)
        lines = extract_meals_from_table(table, date_re)

        if not lines:
            raise RuntimeError("날짜가 있는 표는 찾았지만 메뉴를 추출하지 못했습니다.")

        output_path = get_output_path(file_path, month)
        save_result(lines, output_path)

        messagebox.showinfo("완료", f"저장 완료:\n{output_path}")

    except Exception as e:
        messagebox.showerror("오류", str(e))



if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        try:
            month = get_month_from_filename(file_path)
            date_re = build_date_re(month)

            root = load_section_xml(file_path)
            table = get_meal_table(root, date_re)
            lines = extract_meals_from_table(table, date_re)

            output_path = get_output_path(file_path, month)
            save_result(lines, output_path)

            messagebox.showinfo("완료", f"저장 완료:\n{output_path}")

        except Exception as e:
            messagebox.showerror("오류", str(e))
    else:
        main()