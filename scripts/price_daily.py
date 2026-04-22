
from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

from utils import (
    WorkbookError,
    add_row_below,
    diff_text,
    find_row_by_mill_and_grade,
    find_sheet,
    format_cn_date,
    iter_price_rows,
    load_xlsx,
    locate_header_row,
    read_json,
    save_xlsx,
    safe_filename,
    set_value,
    try_float,
    write_json,
    yuanli_diff_text,
)


@dataclass
class RuntimeConfig:
    sheet_name: Optional[str]
    snapshot_path: str
    output_dir: str
    add_xugang_row_if_missing: bool
    mills_linked_to_sangang: List[str]
    mill_aliases: Dict[str, List[str]]
    grade_aliases: Dict[str, List[str]]
    default_output_prefix: str = "市场价格日报"


def _alias_lookup(value: str, aliases: Dict[str, List[str]]) -> List[str]:
    if value in aliases:
        return aliases[value]
    for canonical, alias_values in aliases.items():
        if value == canonical or value in alias_values:
            return alias_values
    return [value]


def _collect_row_costs(rows):
    previous_price_map = {}
    delivered_spread_map = {}
    for row in rows:
        if row.trader_price is not None:
            previous_price_map[(row.mill, row.grade)] = row.trader_price
        if row.trader_price is not None and row.delivered_price is not None:
            delivered_spread_map[(row.mill, row.grade)] = round(row.delivered_price - row.trader_price)
    return previous_price_map, delivered_spread_map


def _ensure_special_row(ws, header, rows, cfg: RuntimeConfig):
    special = find_row_by_mill_and_grade(rows, _alias_lookup("徐钢，六安，大东海", cfg.mill_aliases), _alias_lookup("45#", cfg.grade_aliases))
    if special or not cfg.add_xugang_row_if_missing:
        return
    linggang = find_row_by_mill_and_grade(rows, _alias_lookup("凌钢", cfg.mill_aliases), _alias_lookup("45#", cfg.grade_aliases))
    if not linggang:
        return
    new_row_idx = add_row_below(ws, linggang.row_index)
    set_value(ws, new_row_idx, header.columns["钢厂"], "徐钢，六安，大东海")
    set_value(ws, new_row_idx, header.columns["钢种"], "45#")
    set_value(ws, new_row_idx, header.columns["贸易商报价"], 3580)
    delivered_col = header.columns.get("泉州地区送到价")
    if delivered_col:
        set_value(ws, new_row_idx, delivered_col, 3600)


def _get_row(rows, cfg: RuntimeConfig, mill_name: str, grade_name: str):
    return find_row_by_mill_and_grade(rows, _alias_lookup(mill_name, cfg.mill_aliases), _alias_lookup(grade_name, cfg.grade_aliases))


def _update_row_value(ws, row, header, new_trader_price: Optional[float], delivered_spread: Optional[float]) -> None:
    if row is None or new_trader_price is None:
        return
    new_trader_price = int(round(new_trader_price))
    set_value(ws, row.row_index, header.columns["贸易商报价"], new_trader_price)
    delivered_col = header.columns.get("泉州地区送到价")
    if delivered_col is not None and delivered_spread is not None:
        set_value(ws, row.row_index, delivered_col, int(round(new_trader_price + delivered_spread)))


def _refresh_change_and_yuanli(ws, rows, header, previous_price_map, cfg: RuntimeConfig) -> None:
    delivered_col = header.columns.get("泉州地区送到价")
    change_col = header.columns.get("涨跌")
    yuanli_diff_col = header.columns.get("和元立差价")
    yuanli_aliases = _alias_lookup("元立", cfg.mill_aliases)

    for row in rows:
        current_price = try_float(ws.cell(row.row_index, header.columns["贸易商报价"]).value)
        if change_col:
            set_value(ws, row.row_index, change_col, diff_text(current_price, previous_price_map.get((row.mill, row.grade))))
        if not yuanli_diff_col or row.mill in yuanli_aliases:
            continue
        current_delivered = try_float(ws.cell(row.row_index, delivered_col).value) if delivered_col else None
        yuanli_row = _get_row(rows, cfg, "元立", row.grade)
        yuanli_delivered = try_float(ws.cell(yuanli_row.row_index, delivered_col).value) if (yuanli_row and delivered_col) else None
        set_value(ws, row.row_index, yuanli_diff_col, yuanli_diff_text(current_delivered, yuanli_delivered))


def _update_remark_text(ws, sangang_45_price: Optional[float]) -> None:
    if sangang_45_price is None:
        return
    suggested = int(round(sangang_45_price - 150))
    pattern = re.compile(r"(元立建议出厂价\s*45#?\s*)\d+")
    replacement = rf"\g<1>{suggested}"
    for row in ws.iter_rows():
        for cell in row:
            text = cell.value
            if isinstance(text, str) and "元立建议出厂价" in text:
                new_text = pattern.sub(replacement, text)
                if new_text == text:
                    new_text = f"{text.rstrip()} 元立建议出厂价45#{suggested}"
                cell.value = new_text
                return


def run(input_json: str) -> str:
    payload = read_json(input_json)
    cfg = RuntimeConfig(**payload["runtime"])
    wb = load_xlsx(payload["template_path"])
    ws = find_sheet(wb, cfg.sheet_name)
    header = locate_header_row(ws)

    original_rows = iter_price_rows(ws, header)
    _ensure_special_row(ws, header, original_rows, cfg)
    rows = iter_price_rows(ws, header)
    previous_price_map, delivered_spread_map = _collect_row_costs(rows)

    changed_prices = payload["inputs"]["changed_prices"]
    report_date = payload["inputs"].get("report_date")
    output_filename = payload["inputs"].get("output_filename") or safe_filename(cfg.default_output_prefix, report_date)

    base_sangang_45 = changed_prices.get("三钢", {}).get("45#")
    if base_sangang_45 is None:
        raise WorkbookError("输入缺少 三钢 / 45# 价格。")

    for linked_mill in cfg.mills_linked_to_sangang:
        row_45 = _get_row(rows, cfg, linked_mill, "45#")
        sangang_45_row = _get_row(rows, cfg, "三钢", "45#")
        if row_45 and sangang_45_row and row_45.trader_price is not None and sangang_45_row.trader_price is not None:
            fixed_spread = round(row_45.trader_price - sangang_45_row.trader_price)
            _update_row_value(ws, row_45, header, base_sangang_45 + fixed_spread, delivered_spread_map.get((row_45.mill, row_45.grade)))

    for mill_name, grade_map in changed_prices.items():
        for grade_name, new_price in grade_map.items():
            row = _get_row(rows, cfg, mill_name, grade_name)
            if not row:
                continue
            _update_row_value(ws, row, header, try_float(new_price), delivered_spread_map.get((row.mill, row.grade)))

    mills_to_sync_40cr = {row.mill for row in rows if row.grade in _alias_lookup("45#", cfg.grade_aliases) or row.grade in _alias_lookup("40Cr", cfg.grade_aliases)}
    for mill_name in mills_to_sync_40cr:
        row_45 = find_row_by_mill_and_grade(rows, [mill_name], _alias_lookup("45#", cfg.grade_aliases))
        row_40 = find_row_by_mill_and_grade(rows, [mill_name], _alias_lookup("40Cr", cfg.grade_aliases))
        if not row_45 or not row_40:
            continue
        old_45 = previous_price_map.get((row_45.mill, row_45.grade))
        old_40 = previous_price_map.get((row_40.mill, row_40.grade))
        new_45 = try_float(ws.cell(row_45.row_index, header.columns["贸易商报价"]).value)
        if old_45 is None or old_40 is None or new_45 is None:
            continue
        fixed_spread = round(old_40 - old_45)
        _update_row_value(ws, row_40, header, new_45 + fixed_spread, delivered_spread_map.get((row_40.mill, row_40.grade)))

    mills_to_sync_195 = {row.mill for row in rows if row.grade in _alias_lookup("195", cfg.grade_aliases) or row.grade in _alias_lookup("08Al", cfg.grade_aliases)}
    for mill_name in mills_to_sync_195:
        row_08 = find_row_by_mill_and_grade(rows, [mill_name], _alias_lookup("08Al", cfg.grade_aliases))
        row_195 = find_row_by_mill_and_grade(rows, [mill_name], _alias_lookup("195", cfg.grade_aliases))
        if not row_08 or not row_195:
            continue
        old_08 = previous_price_map.get((row_08.mill, row_08.grade))
        old_195 = previous_price_map.get((row_195.mill, row_195.grade))
        new_08 = try_float(ws.cell(row_08.row_index, header.columns["贸易商报价"]).value)
        if old_08 is None or old_195 is None or new_08 is None:
            continue
        fixed_spread = round(old_195 - old_08)
        _update_row_value(ws, row_195, header, new_08 + fixed_spread, delivered_spread_map.get((row_195.mill, row_195.grade)))

    _refresh_change_and_yuanli(ws, rows, header, previous_price_map, cfg)
    _update_remark_text(ws, try_float(base_sangang_45))

    snapshot = {
        "last_output": output_filename,
        "last_generated_at": format_cn_date(report_date),
        "template_path": payload["template_path"],
    }
    write_json(cfg.snapshot_path, snapshot)

    output_path = str(Path(cfg.output_dir) / output_filename)
    save_xlsx(wb, output_path)
    return output_path


def build_parser():
    parser = argparse.ArgumentParser(description="根据既有价差规则生成市场价格日报")
    parser.add_argument("--input", required=True, help="运行参数 JSON")
    return parser


if __name__ == "__main__":
    args = build_parser().parse_args()
    path = run(args.input)
    print(path)
