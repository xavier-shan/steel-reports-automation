
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List, Optional

from utils import WorkbookError, find_sheet, load_xlsx, read_json, save_xlsx, safe_filename, write_json


def rotate_list(values: List[str], start_index: int, count: int = 1):
    size = len(values)
    if size == 0:
        raise WorkbookError("客户轮换列表不能为空。")
    result = []
    idx = start_index
    for _ in range(count):
        result.append(values[idx % size])
        idx += 1
    return result, idx % size


def run(input_json: str) -> str:
    payload = read_json(input_json)
    template_path = payload["template_path"]
    runtime = payload["runtime"]
    inputs = payload["inputs"]
    state_path = runtime["rotation_state_path"]

    state = read_json(state_path)
    wb = load_xlsx(template_path)
    ws = find_sheet(wb, runtime.get("sheet_name"))

    cell_map = runtime["cell_map"]

    # 固定复制逻辑：C5 <- F5，C6 <- F6，C7 <- F7
    for left_cell, right_cell in (("C5", "F5"), ("C6", "F6"), ("C7", "F7")):
        ws[left_cell] = ws[right_cell].value

    round_list = state["round_bar_customers"]
    wire_list = state["wire_customers"]

    next_round, next_round_index = rotate_list(round_list, state["round_bar_index"], 1)
    next_wire, next_wire_index = rotate_list(wire_list, state["wire_index"], 2)

    ws["F5"] = next_round[0]
    ws["F6"] = next_wire[0]
    ws["F7"] = next_wire[1]

    # 福建广吉固定不参与轮换，默认保持在配置位置
    fixed_customer_cells = runtime.get("fixed_customer_cells", {})
    for cell_addr, value in fixed_customer_cells.items():
        ws[cell_addr] = value

    # 按输入更新 B 列送到价
    for cell_addr, value in inputs.get("delivered_prices", {}).items():
        ws[cell_addr] = value

    # 允许用户显式更新其他格子
    for cell_addr, value in inputs.get("direct_updates", {}).items():
        ws[cell_addr] = value

    state["round_bar_index"] = next_round_index
    state["wire_index"] = next_wire_index
    write_json(state_path, state)

    output_filename = inputs.get("output_filename") or safe_filename(runtime.get("default_output_prefix", "经理周报"), inputs.get("report_date"))
    output_path = str(Path(runtime["output_dir"]) / output_filename)
    save_xlsx(wb, output_path)
    return output_path


def build_parser():
    parser = argparse.ArgumentParser(description="根据固定轮换规则生成经理周报")
    parser.add_argument("--input", required=True, help="运行参数 JSON")
    return parser


if __name__ == "__main__":
    args = build_parser().parse_args()
    path = run(args.input)
    print(path)
