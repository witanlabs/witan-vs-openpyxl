"""Case 21: author a combo chart with a secondary axis and exact placement.

The task is to create a clustered column + line combo chart:

- Revenue as columns on the primary value axis.
- Margin as a line on the secondary value axis.
- Secondary value axis on the right with percent number format.
- Marker on the line series.
- Two-cell placement from E2 to M18.

Witan authors this directly through the chart model. openpyxl can create a
combo chart, but its authored chart reads back with rough axis/category/anchor
semantics. LibreOffice repairs some axis semantics on save, but not everything.
"""
from __future__ import annotations

import json
import os
import shlex
import shutil
import subprocess
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "combo_chart"
SOFFICE = Path(os.environ.get("SOFFICE", "/opt/homebrew/bin/soffice"))
WITAN_CMD = shlex.split(os.environ.get("WITAN_CMD", "npx witan"))


def run(cmd: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=ROOT,
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )


def resave_with_libreoffice(src: Path) -> Path:
    if not SOFFICE.exists():
        raise RuntimeError(f"LibreOffice not found at {SOFFICE}; set SOFFICE=/path/to/soffice")

    work = OUT / "_lo_work"
    converted = OUT / "_lo_out"
    profile = OUT / "_lo_profile"
    shutil.rmtree(work, ignore_errors=True)
    shutil.rmtree(converted, ignore_errors=True)
    shutil.rmtree(profile, ignore_errors=True)
    work.mkdir(parents=True)
    converted.mkdir(parents=True)
    profile.mkdir(parents=True)

    local = work / src.name
    shutil.copy(src, local)
    run(
        [
            str(SOFFICE),
            "--headless",
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(converted),
            str(local),
        ]
    )
    out = converted / local.name
    if not out.exists():
        raise RuntimeError(f"LibreOffice did not produce {out}")
    return out


def witan_create() -> dict[str, object]:
    dst = OUT / "case21_witan.xlsx"
    code = r'''
await xlsx.addSheet(wb, "Data");
await xlsx.setCells(wb, [
  {address:"Data!A1", value:"Month"}, {address:"Data!B1", value:"Revenue"}, {address:"Data!C1", value:"Margin"},
  {address:"Data!A2", value:"Jan"},   {address:"Data!B2", value:1000},      {address:"Data!C2", value:0.20},
  {address:"Data!A3", value:"Feb"},   {address:"Data!B3", value:1300},      {address:"Data!C3", value:0.25},
  {address:"Data!A4", value:"Mar"},   {address:"Data!B4", value:1150},      {address:"Data!C4", value:0.22},
  {address:"Data!A5", value:"Apr"},   {address:"Data!B5", value:1500},      {address:"Data!C5", value:0.28}
]);
await xlsx.addChart(wb, "Data", {
  name: "Revenue Margin Combo",
  position: { from: { cell: "E2" }, to: { cell: "M18" } },
  groups: [
    { type: "column", axis: "primary", series: [
      { name: { ref: "Data!B1" }, categories: "Data!A2:A5", values: "Data!B2:B5" }
    ] },
    { type: "line", axis: "secondary", series: [
      { name: { ref: "Data!C1" }, categories: "Data!A2:A5", values: "Data!C2:C5", marker: { symbol: "circle" } }
    ] }
  ],
  title: { text: "Revenue and Margin" },
  legend: { position: "right" },
  axes: {
    value: { title: { text: "Revenue" } },
    secondaryValue: { title: { text: "Margin" }, numberFormat: "0%" }
  }
});
return await xlsx.getChart(wb, "Data", "Revenue Margin Combo");
'''
    proc = run([*WITAN_CMD, "xlsx", "exec", str(dst), "--create", "--save", "--code", code])
    return json.loads(proc.stdout)


def openpyxl_create() -> Path:
    dst = OUT / "case21_openpyxl.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    rows = [
        ["Month", "Revenue", "Margin"],
        ["Jan", 1000, 0.20],
        ["Feb", 1300, 0.25],
        ["Mar", 1150, 0.22],
        ["Apr", 1500, 0.28],
    ]
    for row in rows:
        ws.append(row)

    bar = BarChart()
    bar.type = "col"
    bar.title = "Revenue and Margin"
    bar.y_axis.title = "Revenue"
    bar.x_axis.title = "Month"
    bar.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=5))

    line = LineChart()
    line.y_axis.axId = 200
    line.y_axis.title = "Margin"
    line.y_axis.numFmt = "0%"
    line.y_axis.crosses = "max"
    line.add_data(Reference(ws, min_col=3, min_row=1, max_row=5), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=2, max_row=5))

    bar += line
    ws.add_chart(bar, "E2")
    wb.save(dst)
    return dst


def get_chart(path: Path, name: str) -> dict[str, object]:
    proc = run([*WITAN_CMD, "xlsx", "exec", str(path), "--expr", f'xlsx.getChart(wb, "Data", "{name}")'])
    return json.loads(proc.stdout)


def summarize(chart: dict[str, object]) -> dict[str, object]:
    groups = chart.get("groups", [])
    line_group = next((g for g in groups if g.get("type") == "line"), {})
    line_series = line_group.get("series", [{}])[0] if line_group else {}
    return {
        "position": chart.get("position"),
        "groupTypes": [g.get("type") for g in groups],
        "groupAxes": [g.get("axis") for g in groups],
        "categoryPosition": chart.get("axes", {}).get("category", {}).get("position"),
        "secondaryValuePosition": chart.get("axes", {}).get("secondaryValue", {}).get("position"),
        "secondaryValueFormat": chart.get("axes", {}).get("secondaryValue", {}).get("numberFormat"),
        "lineMarker": line_series.get("marker"),
        "lineCategoriesType": line_series.get("categoriesRefType"),
    }


def is_witan_expected(summary: dict[str, object]) -> bool:
    pos = summary["position"]
    return (
        summary["groupTypes"] == ["line", "column"]
        and summary["groupAxes"] == ["secondary", "primary"]
        and summary["secondaryValuePosition"] == "right"
        and summary["secondaryValueFormat"] == "0%"
        and summary["lineMarker"] == {"size": 5, "style": "circle"}
        and pos["from"]["cell"] == "E2"
        and pos["to"]["cell"] == "M18"
    )


def main() -> int:
    shutil.rmtree(OUT, ignore_errors=True)
    OUT.mkdir(parents=True)

    witan_chart = witan_create()
    witan_summary = summarize(witan_chart)

    openpyxl_path = openpyxl_create()
    openpyxl_summary = summarize(get_chart(openpyxl_path, "Chart 1"))

    lo_path = resave_with_libreoffice(openpyxl_path)
    lo_summary = summarize(get_chart(lo_path, "Chart 1"))

    print("21 Combo chart with secondary axis")
    print(f"   witan: {witan_summary}")
    print(f"   openpyxl: {openpyxl_summary}")
    print(f"   openpyxl + LibreOffice: {lo_summary}")
    print()

    unexpected = not is_witan_expected(witan_summary) or is_witan_expected(openpyxl_summary) or is_witan_expected(lo_summary)
    print(f"Summary: {'expected comparison' if not unexpected else 'unexpected result'}")
    return 1 if unexpected else 0


if __name__ == "__main__":
    raise SystemExit(main())
