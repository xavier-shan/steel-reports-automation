
from __future__ import annotations

import argparse
import shutil
import subprocess
import tempfile
from pathlib import Path


def which_or_raise(candidates):
    for candidate in candidates:
        if shutil.which(candidate):
            return shutil.which(candidate)
        if Path(candidate).exists():
            return candidate
    raise FileNotFoundError("未找到 LibreOffice 或 ImageMagick，请先安装。")


def run(xlsx_path: str, output_png: str) -> str:
    xlsx = Path(xlsx_path).resolve()
    output = Path(output_png).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    soffice = which_or_raise([
        "soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ])
    magick = which_or_raise(["magick"])

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(tmpdir), str(xlsx)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        pdf_path = tmpdir / f"{xlsx.stem}.pdf"
        if not pdf_path.exists():
            raise FileNotFoundError("LibreOffice 未产出 PDF，请检查模板打印区域。")
        subprocess.run(
            [magick, "-density", "220", f"{pdf_path}[0]", "-quality", "92", str(output)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    return str(output)


def build_parser():
    parser = argparse.ArgumentParser(description="把 Excel 产物导出成整页 PNG 预览")
    parser.add_argument("--xlsx", required=True)
    parser.add_argument("--png", required=True)
    return parser


if __name__ == "__main__":
    args = build_parser().parse_args()
    print(run(args.xlsx, args.png))
