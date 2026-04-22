
from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from utils import (
    WorkbookError,
    add_row_below,
    build_row_index,
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
    return aliases.get(value, [value])


def _collect_row_costs(rows):
    price_diff_map = {}
    delivered_spread_map = {}
    for row in rows:
        if row.trader_price is not None:
            price_diff_map[(row.mill, row.grade)] = row.trader_price
        if row.trader_price is not None and row.delivered_price is not None:
            delivered_spread_map[(row.mill, row.grade)] = round(row.delivered_price - row.trader_price)
    return price_diff_map, delivered_spread_map


def _ensure_special_row(ws, header, rows, cfg: RuntimeConfig):
    special_mills = cfg.mill_aliases.get("徐钢，六安，大东海", ["徐钢，六安，大东海"])
    special = find_row_by_mill_and_grade(rows, special_mills, cfg.grade_aliases.get("45#", ["45#"]))
    if special or not cfg.add_xugang_row_if_missing:
        return
    linggang = find_row_by_mill_and_grade(rows, cfg.mill_aliases.get("凌钢", ["凌钢"]), cfg.grade_aliases.get("45#", ["45#"]))
    if not linggang:
        return
    new_row_idx = add_row_below(ws, linggang.row_index)
    set_value(ws, new_row_idx, header.columns["钢厂"], "徐钢，六安，大东海")
    set_value(ws, new_row_idx, header.columns["钢种"], "45#")
    set_value(ws, new_row_idx, header.columns["贸易商报价"], 3580)
    delivered_col = header.columns.get("泉州地区送到价")
    if delivered_col:
        set_value(ws, new_row_idx, delivered_col, 3600)


def _get_anchor_price(rows, cfg: RuntimeConfig, mill_name: str, grade_name: str) -> Optional[float]:
    row = find_row_by_mill_and_grade(rows, _alias_lookup(mill_name, cfg.mill_aliases), _alias_lookup(grade_name, cfg.grade_aliases))
    return row.trader_price if row else None


def _update_row_from_price(
    ws,
    row,
    header,
    new_trader_price: Optional[float],
    old_trader_price: Optional[float],
    delivered_spread: Optional[float],
    yuanli_delivered_cost: Optional[float],
):
    if new_trader_price is None or row is None:
        return
    set_value(ws, row.row_index, header.columns["贸易商报价"], new_trader_price)
    new_delivered = None
    delivered_col = header.columns.get("泉州地区送到价")
    if delivered_col and delivered_spread is not None:
        new_delivered = new_trader_price + delivered_spread
        set_value(ws, row.row_index, delivered_col, new_delivered)
    change_col = header.columns.get("涨跌")
    if change_col:
        set_value(ws, row.row_index, change_col, diff_text(new_trader_price, old_trader_price))
    yuanli_diff_col = header.columns.get("和元立差价")
    if yuanli_diff_col and new_delivered is not None and yuanli_delivered_cost is not None:
        set_value(ws, row.row_index, yuanli_diff_col, yuanli_diff_text(new_delivered, yuanli_delivered_cost))


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
        for grade_name in ("45#", "40Cr"):
            target_row = find_row_by_mill_and_grade(rows, _alias_lookup(linked_mill, cfg.mill_aliases), _alias_lookup(grade_name, cfg.grade_aliases))
            if not target_row:
                continue
            anchor_old = _get_anchor_price(rows, cfg, "三钢", grade_name)
            target_old = target_row.trader_price
            if anchor_old is None or target_old is None:
                continue
            fixed_spread = round(target_old - anchor_old)
            target_new = base_sangang_45 + fixed_spread if grade_name == "45#" else None
            if grade_name == "40Cr":
                sangang_40_old = _get_anchor_price(rows, cfg, "三钢", "40Cr")
                if sangang_40_old is not None:
                    sangang_40_new = base_sangang_45 + round(sangang_40_old - anchor_old)
                    target_new = sangang_40_new + fixed_spread
            yuanli_delivered = None
            yuanli_row = find_row_by_mill_and_grade(rows, _alias_lookup("元立", cfg.mill_aliases), _alias_lookup(grade_name, cfg.grade_aliases))
            if yuanli_row and yuanli_row.delivered_price is not None:
                yuanli_delivered = yuanli_row.delivered_price
            _update_row_from_price(
                ws,
                target_row,
                header,
                target_new,
                target_old,
                delivered_spread_map.get((target_row.mill, target_row.grade)),
                yuanli_delivered,
            )

    for mill_name, grade_map in changed_prices.items():
        for grade_name, new_price in grade_map.items():
            target_row = find_row_by_mill_and_grade(rows, _alias_lookup(mill_name, cfg.mill_aliases), _alias_lookup(grade_name, cfg.grade_aliases))
            if not target_row:
                continue
            yuanli_delivered = None
            yuanli_row = find_row_by_mill_and_grade(rows, _alias_lookup("元立", cfg.mill_aliases), _alias_lookup(grade_name, cfg.grade_aliases))
            if yuanli_row and yuanli_row.delivered_price is not None:
                yuanli_delivered = yuanli_row.delivered_price
            _update_row_from_price(
                ws,
                target_row,
                header,
                try_float(new_price),
                previous_price_map.get((target_row.mill, target_row.grade)),
                delivered_spread_map.get((target_row.mill, target_row.grade)),
                yuanli_delivered,
            )

    for mill_name in ("三钢", "新兴铸管", "凌钢", "徐钢，六安，大东海", "元立"):
        row_45 = find_row_by_mill_and_grade(rows, _alias_lookup(mill_name, cfg.mill_aliases), _alias_lookup("45#", cfg.grade_aliases))
        row_40 = find_row_by_mill_and_grade(rows, _alias_lookup(mill_name, cfg.mill_aliases), _alias_lookup("40Cr", cfg.grade_aliases))
        if not row_45 or not row_40:
            continue
        old_45 = previous_price_map.get((row_45.mill, row_45.grade))
        old_40 = previous_price_map.get((row_40.mill, row_40.grade))
        new_45 = try_float(ws.cell(row_45.row_index, header.columns["贸易商报价"]).value)
        if old_45 is None or old_40 is None or new_45 is None:
            continue
        fixed_spread = round(old_40 - old_45)
        new_40 = new_45 + fixed_spread
        yuanli_row_40 = find_row_by_mill_and_grade(rows, _alias_lookup("元立", cfg.mill_aliases), _alias_lookup("40Cr", cfg.grade_aliases))
        yuanli_delivered_40 = yuanli_row_40.delivered_price if yuanli_row_40 and yuanli_row_40.delivered_price is not None else None
        _update_row_from_price(
            ws,
            row_40,
            header,
            new_40,
            old_40,
            delivered_spread_map.get((row_40.mill, row_40.grade)),
            yuanli_delivered_40,
        )

    for mill_name in ("三钢", "元立"):
        row_08 = find_row_by_mill_and_grade(rows, _alias_lookup(mill_name, cfg.mill_aliases), _alias_lookup("08Al", cfg.grade_aliases))
        row_195 = find_row_by_mill_and_grade(rows, _alias_lookup(mill_name, cfg.mill_aliases), _alias_lookup("195", cfg.grade_aliases))
        if not row_08 or not row_195:
            continue
        old_08 = previous_price_map.get((row_08.mill, row_08.grade))
        old_195 = previous_price_map.get((row_195.mill, row_195.grade))
        new_08 = try_float(ws.cell(row_08.row_index, header.columns["贸易商报价"]).value)
        if old_08 is None or old_195 is None or new_08 is None:
            continue
        fixed_spread = round(old_195 - old_08)
        new_195 = new_08 + fixed_spread
        yuanli_row_195 = find_row_by_mill_and_grade(rows, _alias_lookup("元立", cfg.mill_aliases), _alias_lookup("195", cfg.grade_aliases))
        yuanli_delivered_195 = yuanli_row_195.delivered_price if yuanli_row_195 and yuanli_row_195.delivered_price is not None else None
        _update_row_from_price(
            ws,
            row_195,
            header,
            new_195,
            old_195,
            delivered_spread_map.get((row_195.mill, row_195.grade)),
            yuanli_delivered_195,
        )

    remark_col = header.columns.get("备注")
    if remark_col:
        suggested = try_float(changed_prices.get("三钢", {}).get("45#"))
        if suggested is not None:
            suggestion_text = f"元立建议出厂价{int(round(suggested - 150))}"
            for row in rows:
                if row.grade == "45#":
                    current_text = str(ws.cell(row.row_index, remark_col).value or "")
                    if "元立建议出厂价" in current_text or row.mill in _alias_lookup("元立", cfg.mill_aliases):
                        set_value(ws, row.row_index, remark_col, suggestion_text)

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
