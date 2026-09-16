#!/usr/bin/env python3
"""Обновление региональных ипотечных таблиц по новой выгрузке.

Берёт свежескачанный "Статистические ряды_регионы.xlsx" (тот же формат
выгрузки: листы 01_02_00.."01_02_05", "Данные") и переносит новые
месяцы/годы в мастер-файл "Статистические ряды_регионы (с таблицами).xlsx"
(листы кол-во/объем/ДДУ, структура "как у Дани"), не трогая уже внесённое
форматирование (автофильтр, ширины колонок, формулы-помощники A/B на
листе ДДУ).

Что делает автоматически:
  - находит в новой выгрузке колонки (годовые итоги/месяцы), которых ещё
    нет в мастер-файле, и добавляет их в нужную хронологическую позицию
    (обычно это просто дозапись в конец, но если наступил новый год —
    новая колонка "Итого <год>" вставляется перед месячными колонками);
  - сверяет уже существующие колонки на пересмотр значений "задним
    числом" (частая история со статистикой) и печатает предупреждения,
    НЕ перезаписывая их молча;
  - копирует формат (число/ширина) соседней колонки для новых колонок.

Что не делает автоматически (требует ручного разбора):
  - если в новой выгрузке появилась/пропала целая строка (новый регион
    или новая программа) — строки не мержатся, скрипт останавливается
    на первом несовпадении и печатает диагностику;
  - файл не коммитится и не пушится — это ты делаешь руками после
    просмотра отчёта (и, если нужно, самого файла в Excel).

По умолчанию — dry-run: только печатает отчёт. Пиши в файл только с
флагом --apply.

Использование:
    python3 update_regional_tables.py --raw path/to/новая_выгрузка.xlsx
    python3 update_regional_tables.py --raw ... --apply
"""
import argparse
import datetime
import re
import sys
from copy import copy

import openpyxl

MASTER_PATH = "Статистические ряды_регионы (с таблицами).xlsx"

DDU_LABEL_RE = re.compile(r"^Покупка по ДДУ,\s*(шт\.|млн руб\.)$")

# кол-во/объем всегда исключают агрегат "Все программы" и программу
# "Ипотека в отдельных регионах" - так было в оригинале Дани для каждого
# региона без исключений (проверено по всем ~90 субъектам), хотя в сырых
# данных (01_02_01) обе строки присутствуют с ненулевыми значениями.
EXCLUDED_PROGRAMS = {"Все программы", "Ипотека в отдельных регионах"}

# (raw sheet, raw label col, raw dim col, raw data-start col,
#  master sheet, master label col, master extra cols before dim,
#  master dim col, master data-start col)
HEADER_ROW = 4
FIRST_DATA_ROW = 5


def col_letter(idx):
    return openpyxl.utils.get_column_letter(idx)


def normalize_key(v):
    """Кэш-дружелюбный ключ колонки: строка как есть, дата -> (год, месяц)."""
    if isinstance(v, datetime.datetime):
        return ("date", v.year, v.month)
    return ("text", v)


def fmt_key(v):
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m")
    return str(v)


def is_header_row(ws, row, data_start_col):
    return ws.cell(row=row, column=data_start_col).value is None


def build_filtered_rows(ws, label_col, dim_col, data_start_col, keep_dim):
    """Возвращает список (raw_row, label, dim) для строк, которые попадают
    в производный лист: заголовки регионов (dim_col is None) плюс строки,
    прошедшие keep_dim(label, dim_value)."""
    rows = []
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        label = ws.cell(row=r, column=label_col).value
        if label is None and ws.cell(row=r, column=data_start_col).value is None:
            continue  # полностью пустая строка, пропускаем
        dim = ws.cell(row=r, column=dim_col).value if dim_col else None
        if dim is None:
            if ws.cell(row=r, column=data_start_col).value is None:
                rows.append((r, label, None))  # заголовок региона
            continue
        if keep_dim(label, dim):
            rows.append((r, label, dim))
    return rows


def build_ddu_rows(ws, label_col, data_start_col):
    rows = []
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        label = ws.cell(row=r, column=label_col).value
        if label is None:
            continue
        stripped = label.strip()
        has_data = ws.cell(row=r, column=data_start_col).value is not None
        if not has_data:
            rows.append((r, label, None))  # заголовок региона/программы
        elif DDU_LABEL_RE.match(stripped):
            rows.append((r, label, stripped))
    return rows


def master_row_signature(ws, label_col, dim_col):
    """dim_col=None -> лист ДДУ: тип строки берём из самой метки (label),
    как в build_ddu_rows, а не из отдельной колонки."""
    sig = []
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        label = ws.cell(row=r, column=label_col).value
        if label is None:
            continue
        if dim_col is None:
            stripped = label.strip()
            dim = stripped if DDU_LABEL_RE.match(stripped) else None
        else:
            dim = ws.cell(row=r, column=dim_col).value
        sig.append((r, label, dim))
    return sig


def check_row_alignment(name, master_sig, raw_rows):
    """master_sig и raw_rows — списки (row, label, dim). Строки должны
    совпадать по (label, dim) в том же порядке. Возвращает список лишних
    raw-строк в конце (новые регионы/программы) либо бросает ошибку при
    несовпадении в середине."""
    m = [(lbl, dim) for _, lbl, dim in master_sig]
    r = [(lbl, dim) for _, lbl, dim in raw_rows]
    n = min(len(m), len(r))
    for i in range(n):
        if m[i] != r[i]:
            raise SystemExit(
                f"[{name}] Строки не совпадают на позиции {i} "
                f"(мастер row={master_sig[i][0]}): "
                f"в мастере {m[i]!r}, в новой выгрузке {r[i]!r}.\n"
                f"Это похоже на новый/пропавший регион или программу — "
                f"мержить строки автоматически не буду, разберись руками."
            )
    if len(r) > len(m):
        extra = raw_rows[len(m):]
        print(
            f"[{name}] В новой выгрузке {len(extra)} новых строк в конце "
            f"(похоже на новый регион/программу), не добавлены "
            f"автоматически:"
        )
        for row, lbl, dim in extra[:10]:
            print(f"    raw row {row}: {lbl!r} / {dim!r}")
        if len(extra) > 10:
            print(f"    ... и ещё {len(extra) - 10}")
    elif len(r) < len(m):
        print(
            f"[{name}] ВНИМАНИЕ: в новой выгрузке МЕНЬШЕ строк, чем в "
            f"мастере ({len(r)} < {len(m)}) — проверь, тот ли файл дали."
        )
    return raw_rows[:n]  # только строки, для которых есть соответствие в мастере


def read_master_header(ws, header_row, data_start_col):
    keys = []
    for c in range(data_start_col, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        keys.append((c, normalize_key(v), v))
    return keys


def copy_col_style(ws, src_col, dst_col, first_row, last_row):
    for r in range(first_row, last_row + 1):
        src = ws.cell(row=r, column=src_col)
        dst = ws.cell(row=r, column=dst_col)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
    src_letter = col_letter(src_col)
    dst_letter = col_letter(dst_col)
    if src_letter in ws.column_dimensions:
        ws.column_dimensions[dst_letter].width = ws.column_dimensions[src_letter].width


def expand_auto_filter(ws):
    ref = ws.auto_filter.ref
    if not ref:
        return
    start, _, end = ref.partition(":")
    m = re.match(r"([A-Z]+)(\d+)", start)
    m2 = re.match(r"([A-Z]+)(\d+)", end)
    if not (m and m2):
        return
    new_end = f"{col_letter(ws.max_column)}{m2.group(2)}"
    ws.auto_filter.ref = f"{start}:{new_end}"


def sync_sheet(name, raw_wb, master_wb, raw_sheet, label_col, dim_col,
               data_start_col, master_sheet, master_label_col,
               master_dim_col, master_data_start_col, keep_dim, apply_changes,
               revisions_out):
    raw_ws = raw_wb[raw_sheet]
    master_ws = master_wb[master_sheet]

    if raw_sheet == "01_02_03":
        raw_rows = build_ddu_rows(raw_ws, label_col, data_start_col)
    else:
        raw_rows = build_filtered_rows(raw_ws, label_col, dim_col, data_start_col, keep_dim)

    master_sig = master_row_signature(master_ws, master_label_col, master_dim_col)
    aligned_raw_rows = check_row_alignment(name, master_sig, raw_rows)

    master_header = read_master_header(master_ws, HEADER_ROW, master_data_start_col)
    master_keys_in_order = [k for _, k, _ in master_header]
    master_key_to_col = {k: c for c, k, _ in master_header}

    raw_header_cells = []
    for c in range(data_start_col, raw_ws.max_column + 1):
        v = raw_ws.cell(row=HEADER_ROW, column=c).value
        if v is None:
            continue
        raw_header_cells.append((c, normalize_key(v), v))

    pointer = 0
    new_cols_added = []
    for raw_c, key, raw_val in raw_header_cells:
        if pointer < len(master_keys_in_order) and master_keys_in_order[pointer] == key:
            master_col = master_data_start_col + pointer
            # сверка значений на пересмотр
            for i, (raw_row, lbl, dim) in enumerate(aligned_raw_rows):
                master_row = master_sig[i][0]
                raw_val_cell = raw_ws.cell(row=raw_row, column=raw_c).value
                master_val_cell = master_ws.cell(row=master_row, column=master_col).value
                if raw_val_cell is None and master_val_cell is None:
                    continue
                differs = False
                if isinstance(raw_val_cell, (int, float)) and isinstance(master_val_cell, (int, float)):
                    differs = abs(raw_val_cell - master_val_cell) > max(1e-6, abs(master_val_cell) * 1e-9)
                else:
                    differs = raw_val_cell != master_val_cell
                if differs:
                    revisions_out.append(
                        (name, fmt_key(raw_val), lbl, master_val_cell, raw_val_cell, master_row)
                    )
            pointer += 1
            continue

        if key in master_key_to_col:
            print(
                f"[{name}] Колонка {fmt_key(raw_val)!r} стоит не по порядку "
                f"относительно мастер-файла — пропускаю, разберись руками."
            )
            continue

        # новая колонка — вставляем перед текущей позицией pointer
        insert_master_col = master_data_start_col + pointer
        new_cols_added.append((insert_master_col, key, raw_val))
        if apply_changes:
            template_col = insert_master_col - 1 if insert_master_col > master_data_start_col else master_data_start_col
            master_ws.insert_cols(insert_master_col)
            master_ws.cell(row=HEADER_ROW, column=insert_master_col).value = raw_val
            copy_col_style(master_ws, template_col, insert_master_col, HEADER_ROW, master_ws.max_row)
            for i, (raw_row, lbl, dim) in enumerate(aligned_raw_rows):
                master_row = master_sig[i][0]
                master_ws.cell(row=master_row, column=insert_master_col).value = (
                    raw_ws.cell(row=raw_row, column=raw_c).value
                )
        master_keys_in_order.insert(pointer, key)
        master_key_to_col = {k: master_data_start_col + i for i, k in enumerate(master_keys_in_order)}
        pointer += 1

    if apply_changes and new_cols_added:
        expand_auto_filter(master_ws)

    return new_cols_added


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--raw", required=True, help="Путь к новой выгрузке 'Статистические ряды_регионы.xlsx'")
    ap.add_argument("--master", default=MASTER_PATH, help="Путь к мастер-файлу (по умолчанию — файл в репо)")
    ap.add_argument("--apply", action="store_true", help="Записать изменения (по умолчанию — только отчёт)")
    args = ap.parse_args()

    print(f"Читаю новую выгрузку: {args.raw}")
    raw_wb = openpyxl.load_workbook(args.raw, data_only=True)
    print(f"Читаю мастер-файл: {args.master}")
    master_wb = openpyxl.load_workbook(args.master, data_only=False)

    revisions = []
    all_new_cols = {}

    all_new_cols["кол-во"] = sync_sheet(
        "кол-во", raw_wb, master_wb,
        raw_sheet="01_02_01", label_col=1, dim_col=2, data_start_col=3,
        master_sheet="кол-во", master_label_col=1, master_dim_col=3, master_data_start_col=4,
        keep_dim=lambda lbl, dim: dim == "Всего, шт." and lbl.strip() not in EXCLUDED_PROGRAMS,
        apply_changes=args.apply, revisions_out=revisions,
    )
    all_new_cols["объем"] = sync_sheet(
        "объем", raw_wb, master_wb,
        raw_sheet="01_02_01", label_col=1, dim_col=2, data_start_col=3,
        master_sheet="объем", master_label_col=1, master_dim_col=3, master_data_start_col=4,
        keep_dim=lambda lbl, dim: dim == "Всего, млн руб." and lbl.strip() not in EXCLUDED_PROGRAMS,
        apply_changes=args.apply, revisions_out=revisions,
    )
    all_new_cols["ДДУ"] = sync_sheet(
        "ДДУ", raw_wb, master_wb,
        raw_sheet="01_02_03", label_col=1, dim_col=None, data_start_col=2,
        master_sheet="ДДУ", master_label_col=3, master_dim_col=None, master_data_start_col=5,
        keep_dim=None,
        apply_changes=args.apply, revisions_out=revisions,
    )

    print("\n=== ОТЧЁТ ===")
    for name, cols in all_new_cols.items():
        if cols:
            print(f"[{name}] Новые колонки ({len(cols)}): " + ", ".join(fmt_key(v) for _, _, v in cols))
        else:
            print(f"[{name}] Новых колонок нет.")

    if revisions:
        print(f"\nПересмотр значений задним числом ({len(revisions)} ячеек, показаны первые 20):")
        for name, colkey, lbl, old, new, row in revisions[:20]:
            print(f"  [{name}] {colkey} / строка {row} ({lbl!r}): было {old!r} -> стало {new!r}")
    else:
        print("\nПересмотров значений в существующих колонках не найдено.")

    if args.apply:
        master_wb.save(args.master)
        print(f"\nСохранено: {args.master}")
        print("Дальше: посмотри диф в Excel/LibreOffice, потом закоммить и запушь руками.")
    else:
        print("\nЭто был dry-run (без --apply), файл не менялся.")


if __name__ == "__main__":
    main()
