"""Case 42: bubble chart authoring and rendering.

Witan authors a bubble chart with bubble-specific properties and renders all
points, including a negative-size bubble when showNegativeBubbles is true.
openpyxl can create a basic bubble chart, but cannot apply modern chart style
IDs and emits rough anchor/axis semantics. LibreOffice resave/export drops key
bubble-label and size semantics and renders only one bubble in this probe.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import BubbleChart, Reference, Series
from openpyxl.chart.label import DataLabelList


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "bubble_chart"
ASSETS = ROOT / "assets" / "preview_rendering"
SOFFICE = Path(os.environ.get("SOFFICE", "/opt/homebrew/bin/soffice"))
PDFTOPPM = Path(os.environ.get("PDFTOPPM", "/opt/homebrew/bin/pdftoppm"))
XLSX_SERVE = Path(
    os.environ.get(
        "XLSX_SERVE",
        str((ROOT / "../witan-alfred/bin/publish/xlsx-serve").resolve()),
    )
)


def run(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=ROOT,
        check=check,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )


def run_xlsx_serve(args: list[str]) -> Any:
    proc = run([str(XLSX_SERVE), *args, "--json"])
    payload = json.loads(proc.stdout)
    if not payload.get("ok"):
        raise RuntimeError(json.dumps(payload, indent=2))
    return payload.get("result")


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


def libreoffice_pdf_to_png(src: Path, dst: Path) -> None:
    if not SOFFICE.exists():
        raise RuntimeError(f"LibreOffice not found at {SOFFICE}; set SOFFICE=/path/to/soffice")
    if not PDFTOPPM.exists():
        raise RuntimeError(f"pdftoppm not found at {PDFTOPPM}; set PDFTOPPM=/path/to/pdftoppm")

    converted = OUT / "_lo_pdf"
    profile = OUT / "_lo_pdf_profile"
    shutil.rmtree(converted, ignore_errors=True)
    shutil.rmtree(profile, ignore_errors=True)
    converted.mkdir(parents=True)
    profile.mkdir(parents=True)

    run(
        [
            str(SOFFICE),
            "--headless",
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(converted),
            str(src),
        ]
    )
    pdf = converted / f"{src.stem}.pdf"
    if not pdf.exists():
        raise RuntimeError(f"LibreOffice did not produce {pdf}")

    tmp_prefix = dst.with_suffix("")
    run([str(PDFTOPPM), "-png", "-singlefile", "-r", "110", str(pdf), str(tmp_prefix)])
    produced = tmp_prefix.with_suffix(".png")
    if produced != dst:
        produced.replace(dst)


def render_with_witan(src: Path, dst: Path) -> None:
    run([str(XLSX_SERVE), "render", str(src), "-r", "Data!A1:M20", "-o", str(dst), "--dpr", "1"])
    if not dst.exists():
        raise RuntimeError(f"Witan render did not produce {dst}")


def witan_create() -> tuple[Path, dict[str, Any]]:
    dst = OUT / "case42_witan_bubble.xlsx"
    code = r'''
await xlsx.addSheet(wb, "Data");
await xlsx.setCells(wb, [
  {address:"Data!A1", value:"X"}, {address:"Data!B1", value:"Y"},  {address:"Data!C1", value:"Size"},
  {address:"Data!A2", value:1},   {address:"Data!B2", value:10},   {address:"Data!C2", value:25},
  {address:"Data!A3", value:2},   {address:"Data!B3", value:20},   {address:"Data!C3", value:-50},
  {address:"Data!A4", value:3},   {address:"Data!B4", value:15},   {address:"Data!C4", value:100}
]);
await xlsx.addChart(wb, "Data", {
  name: "Bubble Risk",
  position: { from: { cell: "E2" }, to: { cell: "M18" } },
  groups: [{
    type: "bubble",
    bubbleScale: 150,
    showNegativeBubbles: true,
    sizeRepresents: "width",
    series: [{
      name: { text: "Series" },
      xValues: "Data!A2:A4",
      yValues: "Data!B2:B4",
      bubbleSizes: "Data!C2:C4",
      dataLabels: { showBubbleSize: true, position: "right" }
    }]
  }],
  title: { text: "Bubble Risk" },
  legend: { position: "right" },
  roundedCorners: false,
  styleId: 269
});
return await xlsx.getChart(wb, "Data", "Bubble Risk");
'''
    chart = run_xlsx_serve(["exec", str(dst), "--create", "--save", "--code", code])
    return dst, chart


def openpyxl_create() -> tuple[Path, str | None]:
    dst = OUT / "case42_openpyxl_bubble.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in [
        ["X", "Y", "Size"],
        [1, 10, 25],
        [2, 20, -50],
        [3, 15, 100],
    ]:
        ws.append(row)

    chart = BubbleChart()
    chart.title = "Bubble Risk"
    chart.bubbleScale = 150
    chart.showNegBubbles = True
    chart.sizeRepresents = "w"
    chart.roundedCorners = False
    chart.legend.position = "r"
    modern_style_error = None
    try:
        chart.style = 269
    except ValueError as error:
        modern_style_error = str(error)

    x_values = Reference(ws, min_col=1, min_row=2, max_row=4)
    y_values = Reference(ws, min_col=2, min_row=2, max_row=4)
    bubble_sizes = Reference(ws, min_col=3, min_row=2, max_row=4)
    series = Series(values=y_values, xvalues=x_values, zvalues=bubble_sizes, title="Series")
    series.dLbls = DataLabelList(showBubbleSize=True, dLblPos="r")
    chart.series.append(series)
    ws.add_chart(chart, "E2")
    wb.save(dst)
    return dst, modern_style_error


def get_chart(path: Path, name: str) -> dict[str, Any]:
    return run_xlsx_serve(["exec", str(path), "--expr", f'xlsx.getChart(wb, "Data", "{name}")'])


def summarize(chart: dict[str, Any]) -> dict[str, Any]:
    group = (chart.get("groups") or [{}])[0]
    series = (group.get("series") or [{}])[0]
    labels = series.get("dataLabels") or {}
    position = chart.get("position", {})
    return {
        "position": f'{position.get("from", {}).get("cell")}:{position.get("to", {}).get("cell")}',
        "groupType": group.get("type"),
        "bubbleScale": group.get("bubbleScale"),
        "showNegativeBubbles": group.get("showNegativeBubbles"),
        "sizeRepresents": group.get("sizeRepresents"),
        "bubbleSizes": series.get("bubbleSizes"),
        "showBubbleSizeLabel": labels.get("showBubbleSize"),
        "labelPosition": labels.get("position"),
        "categoryAxisPosition": chart.get("axes", {}).get("category", {}).get("position"),
        "styleId": chart.get("styleId"),
        "roundedCorners": chart.get("roundedCorners"),
    }


def is_expected_witan(summary: dict[str, Any]) -> bool:
    return (
        summary["position"] == "E2:M18"
        and summary["groupType"] == "bubble"
        and summary["bubbleScale"] == 150
        and summary["showNegativeBubbles"] is True
        and summary["sizeRepresents"] == "width"
        and summary["showBubbleSizeLabel"] is True
        and summary["labelPosition"] == "right"
        and summary["categoryAxisPosition"] == "bottom"
        and summary["styleId"] == 269
        and summary["roundedCorners"] is False
    )


def main() -> int:
    shutil.rmtree(OUT, ignore_errors=True)
    OUT.mkdir(parents=True)
    ASSETS.mkdir(parents=True, exist_ok=True)

    witan_path, witan_chart = witan_create()
    witan_summary = summarize(witan_chart)

    openpyxl_path, modern_style_error = openpyxl_create()
    openpyxl_summary = summarize(get_chart(openpyxl_path, "Chart 1"))

    lo_path = resave_with_libreoffice(openpyxl_path)
    lo_summary = summarize(get_chart(lo_path, "Chart 1"))

    render_with_witan(witan_path, ASSETS / "case42_witan_bubble.png")
    libreoffice_pdf_to_png(witan_path, ASSETS / "case42_libreoffice_bubble.png")

    print("42 Bubble chart authoring and rendering")
    print(f"   witan: {witan_summary}")
    print(f"   openpyxl: {openpyxl_summary}, modern style error: {modern_style_error!r}")
    print(f"   openpyxl + LibreOffice: {lo_summary}")
    print(f"   witan render: {ASSETS / 'case42_witan_bubble.png'}")
    print(f"   LibreOffice PDF export: {ASSETS / 'case42_libreoffice_bubble.png'}")
    print()

    unexpected = (
        not is_expected_witan(witan_summary)
        or is_expected_witan(openpyxl_summary)
        or is_expected_witan(lo_summary)
        or modern_style_error != "Max value is 48"
    )
    print(f"Summary: {'expected comparison' if not unexpected else 'unexpected result'}")
    return 1 if unexpected else 0


if __name__ == "__main__":
    raise SystemExit(main())
