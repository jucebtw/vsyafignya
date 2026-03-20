#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка docx: часть 1 — исправленный текст из «Титульник Часть 2.doc» (чистое оглавление, без полей Word);
часть 2 — сжатый раздел 2.1–2.12 + заключение + список по «уп 02», ~50 стр. всего, без «воды»;
рисунки и таблицы по методичке (подписи, нумерация).
Оригиналы не перезаписываются.
"""

from __future__ import annotations

import html
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

WORK = Path("/workspace")
SRC_DOC = WORK / "Титульник Часть 2.doc"
SRC_UP = WORK / "уп 02 2 часть  (2).docx"
OUT_DOCX = WORK / "Титульник Часть 2 — обновлена только 2-я часть.docx"
TEMPLATE = WORK / "уп 02 2 часть  (2).docx"

SPLIT_MARKER_BODY = "\x0c2. Индивидуальное задание"

# Целевой объём всего документа ~50 стр. (оценка ~65–75 тыс. знаков суммарно)
TARGET_TOTAL_CHARS = 78000


def decode_doc_utf16(path: Path) -> str:
    return path.read_bytes().decode("utf-16-le", errors="ignore")


def parse_toc_entry(line: str) -> str | None:
    if "HYPERLINK" not in line or "\x14" not in line:
        return None
    m = re.search(r"\x01\x14([^\t]+)\t([^\t\x13]+)", line)
    if m:
        a, b = m.group(1).strip(), m.group(2).strip()
        if not b.startswith("PAGEREF") and len(b) > 1:
            return f"{a} {b}"
    m = re.search(r"\x01\x14([^\t\x13]+)\t", line)
    if m:
        return m.group(1).strip()
    return None


def build_part1_lines() -> list[str]:
    """Часть 1: титул, чистое оглавление, текст до раздела 2 без битых полей."""
    text = decode_doc_utf16(SRC_DOC)
    idx = text.find(SPLIT_MARKER_BODY)
    if idx == -1:
        raise RuntimeError("Не найден маркер «2. Индивидуальное задание» в теле .doc")
    part1_raw = text[:idx]
    lines = [ln.rstrip() for ln in re.split(r"[\r\n]+", part1_raw) if ln.strip()]

    out: list[str] = []
    i = 0
    while i < len(lines):
        ln = lines[i]
        ln_clean = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", ln)

        if "КГБПОУ" not in ln_clean and len(out) == 0:
            i += 1
            continue

        if ln_clean == "Содержание":
            out.append("Содержание")
            i += 1
            while i < len(lines) and ("HYPERLINK" in lines[i] or "TOC" in lines[i]):
                t = parse_toc_entry(lines[i])
                if t:
                    out.append(t)
                i += 1
            continue

        if "HYPERLINK" in ln_clean:
            t = parse_toc_entry(ln_clean)
            if t and (not out or out[-1] != t):
                out.append(t)
            i += 1
            continue

        if ln_clean in ("\x15",):
            i += 1
            continue

        out.append(ln_clean)
        i += 1

    return [x for x in out if x.strip()]


def extract_up02_plain() -> str:
    with zipfile.ZipFile(SRC_UP, "r") as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    chunks: list[str] = []
    for t in root.findall(".//w:t", ns):
        if t.text:
            chunks.append(t.text)
    return "\n".join(chunks)


def split_up02_sections(text: str) -> dict[str, str]:
    ts = text.find("Теоретическая часть\nАнализ существующих")
    ind = text.find("Индивидуальное задание\nФормулировка требований")
    conc = text.find("ЗАКЛЮЧЕНИЕ\nВ ходе выполнения")
    bib = text.find("10 Essential Best Practices")
    if -1 in (ts, ind, conc, bib):
        raise RuntimeError("Не найдены границы разделов уп 02")
    return {
        "theory": text[ts:ind],
        "ind": text[ind:conc],
        "conc": text[conc:bib],
        "bib": "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ\n" + text[bib:],
    }


def extract_between(s: str, start: str, end: str) -> str:
    if not s:
        return ""
    i = s.find(start)
    if i == -1:
        return ""
    i += len(start)
    if not end:
        return s[i:].strip()
    j = s.find(end, i)
    if j == -1:
        return s[i:].strip()
    return s[i:j].strip()


def shorten_text(s: str, max_chars: int) -> str:
    s = re.sub(r"\n{3,}", "\n\n", s).strip()
    if len(s) <= max_chars:
        return s
    cut = s[: max_chars + 1]
    last = max(cut.rfind("."), cut.rfind("!"), cut.rfind("?"))
    if last > max_chars // 2:
        return cut[: last + 1].strip()
    return cut.rsplit(" ", 1)[0].strip() + "…"


def wrap_paragraphs(s: str) -> list[str]:
    return [p.strip() for p in re.split(r"\n\s*\n", s) if p.strip()]


def build_part2_lines_compact(up: dict[str, str]) -> list[str]:
    theory, ind = up["theory"], up["ind"]
    conc, bib = up["conc"], up["bib"]

    def sec(title: str, body: str, max_c: int = 5200) -> list[str]:
        body = shorten_text(body, max_c)
        return [title, ""] + wrap_paragraphs(body) + [""]

    lines: list[str] = ["2. Индивидуальное задание", ""]

    lines += sec(
        "2.1 Формулировка требований к программному изделию, моделирование с помощью диаграмм",
        extract_between(ind, "Формулировка требований", "Идентификация элементов"),
        4000,
    )
    lines += sec(
        "2.2 Формулировка функциональных требований к программному изделию, использование методов предпроектного обследования",
        extract_between(ind, "Функциональные требования", "Моделирование системы")
        + "\n"
        + extract_between(ind, "Нефункциональные требования", "Ограничения проекта"),
        3800,
    )
    lines += sec(
        "2.3 Построение информационно-логической схемы данных, формулировка требований по информационному обеспечению программного изделия",
        extract_between(theory, "Проектирование и разработка базы данных", "1.5")
        + "\n"
        + extract_between(ind, "Элементы входной информации", "Рассмотрение программного"),
        4200,
    )
    lines += sec(
        "2.4 Создание функциональной модели предметной области, диаграмм потоков данных (процессов) DFD, диаграмм классов",
        extract_between(theory, "Архитектура", "Проектирование и разработка базы данных")
        + "\n"
        + extract_between(ind, "Моделирование системы", "Идентификация элементов"),
        4500,
    )
    lines += sec(
        "2.5 Формулирование проектных решений по функциональной и информационной структуре разрабатываемого программного изделия, обоснование использования инструментальных средств разработки программного изделия, описание функций и параметров программных средств",
        extract_between(theory, "Средства разработки системы", "Архитектура")
        + "\n"
        + extract_between(theory, "1.5", "Индивидуальное задание"),
        4000,
    )
    lines += sec(
        "2.6 Идентификация элементов и описание свойств входной и выходной информации, алгоритма решения задачи, составление контрольного примера",
        extract_between(ind, "Идентификация элементов", "Рассмотрение программного"),
        4000,
    )
    lines += sec(
        "2.7 Выделение объектов-сущностей, выстраивание инфологической модели задачи",
        extract_between(theory, "Логическая структура базы данных", "Тестирование базы данных"),
        3800,
    )
    lines += sec(
        "2.8 Описание шагов алгоритма в неформальной (словесной) форме",
        extract_between(ind, "Описание шагов алгоритма", "Выполнение резервирования"),
        3200,
    )
    lines += sec(
        "2.9 Выполнение анализа информационных ресурсов и категорирование их конфиденциальности, анализ основных угроз защищённости данных, определение комплекса мер нейтрализации угрозы безопасности данных",
        extract_between(ind, "Нефункциональные требования", "Ограничения проекта")
        + "\n"
        + "Классификация данных: справочники, остатки, персональные данные, журналы аудита. Угрозы: несанкционированный доступ, SQL-инъекции, потеря резервной копии. Меры: роли, параметризованные запросы, шифрование бэкапов, журналирование.",
        3500,
    )
    lines += sec(
        "2.10 Выполнение резервирования и восстановления данных",
        extract_between(ind, "Выполнение резервирования", "Создание документации"),
        3200,
    )
    lines += sec(
        "2.11 Разделение большой задачи на модули, выделение функции каждого модуля, их отлаживания, выполнение комплексной отладки задачи",
        extract_between(ind, "Структура проекта", "Инструкция по запуску"),
        3800,
    )
    lines += sec(
        "2.12 Создание документации программиста",
        extract_between(ind, "Создание документации программиста", "ЗАКЛЮЧЕНИЕ"),
        3500,
    )

    lines.append("ЗАКЛЮЧЕНИЕ")
    lines.append("")
    lines.extend(wrap_paragraphs(shorten_text(extract_between(conc, "ЗАКЛЮЧЕНИЕ", "") or conc, 3500)))
    lines.append("")
    lines.append("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    lines.append("")
    lines.extend(wrap_paragraphs(shorten_text(bib, 12000)))

    # Подрезка общего объёма
    joined = "\n".join(lines)
    if len(joined) > TARGET_TOTAL_CHARS:
        factor = TARGET_TOTAL_CHARS / len(joined)
        lines = [shorten_text(x, max(80, int(len(x) * factor))) if len(x) > 200 else x for x in lines]

    return lines


def extract_figure_blocks_from_template() -> list[str]:
    """Параграф с рисунком + следующий абзац с подписью «Рисунок …»."""
    with zipfile.ZipFile(TEMPLATE, "r") as z:
        xml = z.read("word/document.xml").decode("utf-8")
    paras: list[str] = []
    pos = 0
    while True:
        s = xml.find("<w:p", pos)
        if s == -1:
            break
        e = xml.find("</w:p>", s)
        if e == -1:
            break
        paras.append(xml[s : e + 6])
        pos = e + 6
    blocks: list[str] = []
    for j, p in enumerate(paras):
        if "<w:drawing" not in p:
            continue
        cap = ""
        if j + 1 < len(paras) and "Рисунок" in paras[j + 1]:
            cap = paras[j + 1]
        blocks.append(p + cap)
        if len(blocks) >= 8:
            break
    return blocks


def table_labor_xml() -> str:
    """Таблица 1 — оценка трудозатрат (по методичке: название над таблицей слева)."""
    rows = [
        ["Этап", "Содержание", "Часы"],
        ["Анализ предметной области", "Изучение процессов, аналогов", "16"],
        ["Архитектура", "Структура ИС", "14"],
        ["БД", "Логическая и физическая модель", "18"],
        ["Разработка", "Логика и интерфейс", "40"],
        ["Тестирование", "Проверка системы и БД", "12"],
        ["Документация", "Пояснения и инструкции", "10"],
        ["Итого", "", "110"],
    ]
    return table_with_caption(
        "Таблица 1 – Оценка трудозатрат на разработку системы",
        rows,
    )


def table_compare_xml() -> str:
    """Таблица 2 — сравнение типов систем."""
    rows = [
        ["Тип", "Плюсы", "Минусы"],
        ["ERP", "Интеграция с финансами, единое пространство", "Стоимость, сложность внедрения"],
        ["WMS", "Складская детализация, адресное хранение", "Нужна интеграция с ERP"],
        ["Специализированная", "Адаптация под процесс", "Поддержка своими силами"],
    ]
    return table_with_caption(
        "Таблица 2 – Сравнение классов программных решений для склада",
        rows,
    )


def table_with_caption(caption: str, rows: list[list[str]]) -> str:
    esc_cap = html.escape(caption)
    cap = (
        f"<w:p><w:pPr><w:spacing w:before='120' w:line='360' w:lineRule='auto'/>"
        f"<w:ind w:firstLine='0'/><w:jc w:val='left'/></w:pPr>"
        f"<w:r><w:rPr><w:rFonts w:ascii='Times New Roman' w:hAnsi='Times New Roman'/>"
        f"<w:b/><w:sz w:val='28'/></w:rPr><w:t xml:space='preserve'>{esc_cap}</w:t></w:r></w:p>"
    )
    tbl_rows = []
    for row in rows:
        cells = []
        for cell in row:
            esc = html.escape(cell)
            cells.append(
                f"<w:tc><w:tcPr><w:tcW w:w='3000' w:type='dxa'/></w:tcPr>"
                f"<w:p><w:pPr><w:spacing w:line='360' w:lineRule='auto'/><w:jc w:val='both'/>"
                f"<w:ind w:firstLine='0'/></w:pPr>"
                f"<w:r><w:rPr><w:rFonts w:ascii='Times New Roman' w:hAnsi='Times New Roman'/>"
                f"<w:sz w:val='28'/></w:rPr><w:t xml:space='preserve'>{esc}</w:t></w:r></w:p></w:tc>"
            )
        tbl_rows.append(f"<w:tr>{''.join(cells)}</w:tr>")
    tbl = (
        "<w:tbl><w:tblPr><w:tblW w:w='5000' w:type='pct'/><w:tblBorders>"
        "<w:top w:val='single' w:sz='4'/><w:left w:val='single' w:sz='4'/>"
        "<w:bottom w:val='single' w:sz='4'/><w:right w:val='single' w:sz='4'/>"
        "<w:insideH w:val='single' w:sz='4'/><w:insideV w:val='single' w:sz='4'/>"
        "</w:tblBorders></w:tblPr>"
        + "".join(tbl_rows)
        + "</w:tbl>"
    )
    return cap + tbl


def clean_text_for_word(s: str) -> str:
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\uFFFD]", "", s)
    return s.strip()


def paragraph_xml(text: str, bold: bool = False, heading: bool = False) -> str:
    text = clean_text_for_word(text)
    if not text:
        return ""
    esc = html.escape(text)
    base_rfonts = (
        "<w:rFonts w:ascii='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman'/>"
        "<w:sz w:val='28'/><w:szCs w:val='28'/>"
    )
    if bold or heading:
        r_pr = "<w:rPr><w:b/><w:bCs/>" + base_rfonts + "</w:rPr>"
    else:
        r_pr = "<w:rPr>" + base_rfonts + "</w:rPr>"
    p_pr = (
        "<w:pPr><w:spacing w:before='240' w:after='120' w:line='360' w:lineRule='auto'/>"
        "<w:ind w:firstLine='708'/><w:jc w:val='both'/></w:pPr>"
        if heading
        else "<w:pPr><w:spacing w:line='360' w:lineRule='auto'/>"
        "<w:ind w:firstLine='708'/><w:jc w:val='both'/></w:pPr>"
    )
    return f"<w:p>{p_pr}<w:r>{r_pr}<w:t xml:space='preserve'>{esc}</w:t></w:r></w:p>"


def page_break_xml() -> str:
    return "<w:p><w:r><w:br w:type='page'/></w:r></w:p>"


def build_document_xml(
    part1_lines: list[str],
    part2_segments: list,
    doc_open: str,
    sect_pr: str,
) -> bytes:
    parts: list[str] = [doc_open, "<w:body>"]

    for line in part1_lines:
        cleaned = clean_text_for_word(line)
        if not cleaned:
            continue
        bold_title = cleaned in (
            "Содержание",
            "Введение",
            "Сведения о практике",
            "1 Основная часть",
        ) or re.match(
            r"^(1\.\d|2\.\d|Заключение|Список)", cleaned, re.I
        )
        if re.match(r"^1\.6\.\d", cleaned):
            parts.append(paragraph_xml(cleaned, bold=True))
        elif bold_title and len(cleaned) < 120:
            parts.append(paragraph_xml(cleaned, bold=True))
        else:
            parts.append(paragraph_xml(cleaned))

    parts.append(page_break_xml())

    for item in part2_segments:
        if item[0] == "text":
            t, b, h = item[1], item[2], item[3]
            parts.append(paragraph_xml(t, bold=b, heading=h))
        elif item[0] == "raw":
            parts.append(item[1])
        elif item[0] == "tbl":
            parts.append(item[1])

    parts.append(sect_pr)
    parts.append("</w:body></w:document>")
    return "".join(parts).encode("utf-8")


def build_segments_simple(lines: list[str], figures: list[str]) -> list:
    """Линейная сборка: текст по строкам, таблицы и рисунки в фиксированных точках."""
    segments: list = []
    inserted_tbl = False
    inserted_tbl2 = False
    inserted_fig1 = False
    inserted_fig2 = False
    for ln in lines:
        if not ln.strip():
            continue
        st = ln.strip()
        if st in ("2. Индивидуальное задание", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"):
            segments.append(("text", st, True, True))
        elif re.match(r"^2\.\d{1,2}\s", st):
            segments.append(("text", st, True, True))
            if st.startswith("2.3 ") and not inserted_tbl:
                segments.append(("tbl", table_labor_xml()))
                inserted_tbl = True
            if st.startswith("2.4 ") and not inserted_tbl2:
                segments.append(("tbl", table_compare_xml()))
                inserted_tbl2 = True
            if st.startswith("2.4 ") and not inserted_fig1:
                segments.append(
                    (
                        "text",
                        "Рисунки 1–4 — диаграммы и схемы по материалам ПМ.02 (см. пояснения в тексте раздела).",
                        False,
                        False,
                    )
                )
                for j in range(min(4, len(figures))):
                    segments.append(("raw", figures[j]))
                inserted_fig1 = True
        else:
            segments.append(("text", st, False, False))

        if (
            not inserted_fig2
            and len(figures) > 4
            and ("Рассмотрение программного" in st or "Основной функционал системы" in st)
        ):
            segments.append(
                (
                    "text",
                    "Рисунки 5–8 — экранные формы приложения (Avalonia UI).",
                    False,
                    False,
                )
            )
            for j in range(4, min(8, len(figures))):
                segments.append(("raw", figures[j]))
            inserted_fig2 = True

    if len(figures) > 4 and not inserted_fig2:
        segments.append(
            (
                "text",
                "Рисунки 5–8 — экранные формы автоматизированной системы.",
                False,
                False,
            )
        )
        for j in range(4, min(8, len(figures))):
            segments.append(("raw", figures[j]))
    return segments


def main() -> None:
    part1_lines = build_part1_lines()
    figures = extract_figure_blocks_from_template()
    if len(figures) < 8:
        print(f"Предупреждение: найдено рисунков {len(figures)}, ожидалось 8")

    up = split_up02_sections(extract_up02_plain())
    lines = build_part2_lines_compact(up)
    segments = build_segments_simple(lines, figures)

    with zipfile.ZipFile(TEMPLATE, "r") as z:
        xml = z.read("word/document.xml").decode("utf-8")
    doc_open = xml[: xml.find("<w:body")]
    sect_pr = xml[xml.rfind("<w:sectPr") : xml.find("</w:body>")]

    doc_bytes = build_document_xml(part1_lines, segments, doc_open, sect_pr)

    shutil.copyfile(TEMPLATE, OUT_DOCX)
    tmp = str(OUT_DOCX) + ".tmp"
    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, doc_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))
    Path(tmp).replace(OUT_DOCX)

    p2_len = sum(len(x[1]) for x in segments if x[0] == "text")
    print(f"Готово: {OUT_DOCX} (оценка знаков раздела 2+табл.+подписи: ~{p2_len})")


if __name__ == "__main__":
    main()
