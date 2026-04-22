
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List

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

    copy_pairs = [
        (cell_map["round_current"], cell_map["round_next"]),
        (cell_map["wire_current_1"], cell_map["wire_next_1"]),
        (cell_map["wire_current_2"], cell_map["wire_next_2"]),
    ]
    for left_cell, right_cell in copy_pairs:
        ws[left_cell] = ws[right_cell].value

    round_list = state["round_bar_customers"]
    wire_list = state["wire_customers"]

    next_round, next_round_index = rotate_list(round_list, state["round_bar_index"], 1)
    next_wire, next_wire_index = rotate_list(wire_list, state["wire_index"], 2)

    ws[cell_map["round_next"]] = next_round[0]
    ws[cell_map["wire_next_1"]] = next_wire[0]
    ws[cell_map["wire_next_2"]] = next_wire[1]

    for cell_addr, value in runtime.get("fixed_customer_cells", {}).items():
        ws[cell_addr] = value

    for cell_addr, value in inputs.get("delivered_prices", {}).items():
        ws[cell_addr] = value

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
