"""Case 25: author a smooth scatter chart with markers and exact placement.

The task is to create a smooth XY scatter chart with circle markers:

- X values from Data!A2:A4 and Y values from Data!B2:B4.
- Smooth line with visible circle markers.
- Legend at the bottom.
- Witan chart style 240.
- Two-cell placement from E2 to M18.

openpyxl can create a scatter chart when the series marker is set explicitly,
but the authored chart still reads back with rough anchor/style/axis semantics.
LibreOffice repairs the axis interpretation and applies its own default style,
but it does not make the chart equivalent to the Witan-authored spec.
"""
from __future__ import annotations

import json
import os
import shlex
import shutil
import subprocess
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.marker import Marker


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "scatter_chart"
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


def witan_create() -> dict[str, Any]:
    dst = OUT / "case25_witan.xlsx"
    code = r'''
await xlsx.addSheet(wb, "Data");
await xlsx.setCells(wb, [
  {address:"Data!A1", value:"X"}, {address:"Data!B1", value:"Y"},
  {address:"Data!A2", value:1},   {address:"Data!B2", value:3},
  {address:"Data!A3", value:2},   {address:"Data!B3", value:5},
  {address:"Data!A4", value:3},   {address:"Data!B4", value:4}
]);
await xlsx.addChart(wb, "Data", {
  name: "Response Curve",
  position: { from: { cell: "E2" }, to: { cell: "M18" } },
  groups: [{
    type: "scatter",
    scatterStyle: "smoothMarker",
    series: [{
      name: { ref: "Data!B1" },
      xValues: "Data!A2:A4",
      values: "Data!B2:B4",
      marker: { symbol: "circle" }
    }]
  }],
  title: { text: "Response Curve" },
  legend: { position: "bottom" },
  axes: {
    category: { title: { text: "Input" } },
    value: { title: { text: "Output" } }
  },
  styleId: 240
});
return await xlsx.getChart(wb, "Data", "Response Curve");
'''
    proc = run([*WITAN_CMD, "xlsx", "exec", str(dst), "--create", "--save", "--code", code])
    return json.loads(proc.stdout)


def openpyxl_create() -> Path:
    dst = OUT / "case25_openpyxl.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in [
        ["X", "Y"],
        [1, 3],
        [2, 5],
        [3, 4],
    ]:
        ws.append(row)

    chart = ScatterChart()
    chart.title = "Response Curve"
    chart.scatterStyle = "smoothMarker"
    chart.legend.position = "b"
    chart.x_axis.title = "Input"
    chart.y_axis.title = "Output"

    series = Series(
        Reference(ws, min_col=2, min_row=1, max_row=4),
        Reference(ws, min_col=1, min_row=2, max_row=4),
        title_from_data=True,
    )
    series.marker = Marker("circle")
    chart.series.append(series)
    ws.add_chart(chart, "E2")
    wb.save(dst)
    return dst


def get_chart(path: Path, name: str) -> dict[str, Any]:
    proc = run([*WITAN_CMD, "xlsx", "exec", str(path), "--expr", f'xlsx.getChart(wb, "Data", "{name}")'])
    return json.loads(proc.stdout)


def summarize(chart: dict[str, Any]) -> dict[str, Any]:
    groups = chart.get("groups", [])
    group = groups[0] if groups else {}
    series = (group.get("series") or [{}])[0]
    position = chart.get("position", {})
    return {
        "position": f'{position.get("from", {}).get("cell")}:{position.get("to", {}).get("cell")}',
        "groupType": group.get("type"),
        "scatterStyle": group.get("scatterStyle"),
        "smooth": series.get("smooth"),
        "marker": series.get("marker"),
        "xValues": series.get("xValues"),
        "yValues": series.get("yValues"),
        "legendPosition": chart.get("legend", {}).get("position"),
        "categoryAxisPosition": chart.get("axes", {}).get("category", {}).get("position"),
        "styleId": chart.get("styleId"),
    }


def is_witan_expected(summary: dict[str, Any]) -> bool:
    return (
        summary["position"] == "E2:M18"
        and summary["groupType"] == "scatter"
        and summary["scatterStyle"] == "smoothMarker"
        and summary["smooth"] is True
        and summary["marker"] == {"size": 5, "style": "circle"}
        and summary["legendPosition"] == "bottom"
        and summary["categoryAxisPosition"] == "bottom"
        and summary["styleId"] == 240
    )


def main() -> int:
    shutil.rmtree(OUT, ignore_errors=True)
    OUT.mkdir(parents=True)

    witan_summary = summarize(witan_create())

    openpyxl_path = openpyxl_create()
    openpyxl_summary = summarize(get_chart(openpyxl_path, "Chart 1"))

    lo_path = resave_with_libreoffice(openpyxl_path)
    lo_summary = summarize(get_chart(lo_path, "Chart 1"))

    print("25 Smooth scatter chart with markers")
    print(f"   witan: {witan_summary}")
    print(f"   openpyxl: {openpyxl_summary}")
    print(f"   openpyxl + LibreOffice: {lo_summary}")
    print()

    unexpected = not is_witan_expected(witan_summary) or is_witan_expected(openpyxl_summary) or is_witan_expected(lo_summary)
    print(f"Summary: {'expected comparison' if not unexpected else 'unexpected result'}")
    return 1 if unexpected else 0


if __name__ == "__main__":
    raise SystemExit(main())
