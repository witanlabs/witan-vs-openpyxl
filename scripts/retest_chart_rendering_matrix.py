"""Cases 43-51: chart rendering matrix split by chart family/variation.

This is a visual comparison between Witan's direct worksheet-range PNG render
and LibreOffice Calc's PDF export rasterized to PNG. The fixtures come from the
local Witan Alfred chart render matrix.

openpyxl is not in the rendering path because it has no native chart/render API.
"""
from __future__ import annotations

import os
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
WITAN_ALFRED = (ROOT / "../witan-alfred").resolve()
OUT = ROOT / "outputs" / "chart_rendering_matrix"
ASSETS = ROOT / "assets" / "chart_rendering"
SOFFICE = Path(os.environ.get("SOFFICE", "/opt/homebrew/bin/soffice"))
PDFTOPPM = Path(os.environ.get("PDFTOPPM", "/opt/homebrew/bin/pdftoppm"))
XLSX_SERVE = Path(
    os.environ.get(
        "XLSX_SERVE",
        str((WITAN_ALFRED / "bin/publish/xlsx-serve").resolve()),
    )
)


@dataclass(frozen=True)
class RenderCase:
    number: int
    key: str
    family: str
    variation: str
    source: Path
    range_address: str


CASES = [
    RenderCase(
        43,
        "bar_negative_clustered",
        "bar/column",
        "clustered columns with negative values",
        WITAN_ALFRED / "fixtures/charts/bar-test.xlsx",
        "NegClustered!A1:R25",
    ),
    RenderCase(
        43,
        "bar_gradient_fill",
        "bar/column",
        "gradient-filled columns",
        WITAN_ALFRED / "fixtures/charts/bar-test.xlsx",
        "GradientFill!A1:R25",
    ),
    RenderCase(
        44,
        "line_gap_span",
        "line",
        "missing values displayed as span",
        WITAN_ALFRED / "fixtures/charts/line-test.xlsx",
        "GapAsSpan!A1:R25",
    ),
    RenderCase(
        44,
        "line_hilow_updown",
        "line",
        "hi-low and up/down bars",
        WITAN_ALFRED / "fixtures/charts/line-test.xlsx",
        "HiLowUpDown!A1:R25",
    ),
    RenderCase(
        45,
        "area_negative_stacked",
        "area",
        "stacked area with negative values",
        WITAN_ALFRED / "fixtures/charts/area-test.xlsx",
        "NegativeStacked!A1:R25",
    ),
    RenderCase(
        46,
        "pie_explode_mixed",
        "pie/doughnut",
        "mixed exploded pie slices",
        WITAN_ALFRED / "fixtures/charts/pie-test.xlsx",
        "ExplodeMixed!A1:R25",
    ),
    RenderCase(
        46,
        "doughnut_two_ring",
        "pie/doughnut",
        "two-ring doughnut",
        WITAN_ALFRED / "fixtures/charts/pie-test.xlsx",
        "TwoRing!A1:R25",
    ),
    RenderCase(
        47,
        "scatter_trendline_power",
        "scatter",
        "scatter with power trendline",
        WITAN_ALFRED / "fixtures/charts/trendline-test.xlsx",
        "ScatterPower!A1:R25",
    ),
    RenderCase(
        48,
        "error_bars_custom",
        "error bars",
        "custom asymmetric error bars",
        WITAN_ALFRED / "fixtures/charts/error-bars-test.xlsx",
        "CustomAsymmetric!A1:R25",
    ),
    RenderCase(
        49,
        "stock_ohlc_multilevel",
        "stock",
        "OHLC with multi-level categories",
        WITAN_ALFRED / "fixtures/charts/stock-test.xlsx",
        "OHLCMultiLevel!A1:R25",
    ),
    RenderCase(
        50,
        "layout_secondary_axis",
        "combo/layout",
        "secondary-axis layout",
        WITAN_ALFRED / "fixtures/charts/layout-test.xlsx",
        "SecondaryAxis!A1:R25",
    ),
    RenderCase(
        51,
        "axis_display_units",
        "axis properties",
        "value axis display units in millions",
        WITAN_ALFRED / "fixtures/charts/axis-test.xlsx",
        "DisplayUnitsMillions!A1:R25",
    ),
]


def run(cmd: list[str], timeout: int | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=ROOT,
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        timeout=timeout,
    )


def sheet_name(range_address: str) -> str:
    if range_address.startswith("'"):
        end = range_address.index("'!")
        return range_address[1:end].replace("''", "'")
    return range_address.split("!", 1)[0]


def isolate_sheet(case: RenderCase) -> Path:
    target_sheet = sheet_name(case.range_address)
    dst = OUT / f"{case.key}.xlsx"
    shutil.copy(case.source, dst)
    code = f"""
const target = {target_sheet!r};
const sheets = await xlsx.listSheets(wb);
for (const sheet of sheets) {{
  if (sheet.sheet !== target) {{
    await xlsx.deleteSheet(wb, sheet.sheet);
  }}
}}
return await xlsx.listSheets(wb);
"""
    run([str(XLSX_SERVE), "exec", str(dst), "--save", "--code", code, "--json"], timeout=30)
    return dst


def render_with_witan(src: Path, range_address: str, dst: Path) -> None:
    run([str(XLSX_SERVE), "render", str(src), "-r", range_address, "-o", str(dst), "--dpr", "1"], timeout=30)
    if not dst.exists():
        raise RuntimeError(f"Witan render did not produce {dst}")


def libreoffice_pdf_to_png(src: Path, dst: Path) -> None:
    if not SOFFICE.exists():
        raise RuntimeError(f"LibreOffice not found at {SOFFICE}; set SOFFICE=/path/to/soffice")
    if not PDFTOPPM.exists():
        raise RuntimeError(f"pdftoppm not found at {PDFTOPPM}; set PDFTOPPM=/path/to/pdftoppm")

    converted = OUT / "_lo_pdf" / src.stem
    profile = OUT / "_lo_profiles" / src.stem
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
        ],
        timeout=30,
    )
    pdf = converted / f"{src.stem}.pdf"
    if not pdf.exists():
        raise RuntimeError(f"LibreOffice did not produce {pdf}")

    tmp_prefix = dst.with_suffix("")
    run([str(PDFTOPPM), "-png", "-singlefile", "-r", "110", str(pdf), str(tmp_prefix)], timeout=30)
    produced = tmp_prefix.with_suffix(".png")
    if produced != dst:
        produced.replace(dst)


def main() -> int:
    shutil.rmtree(OUT, ignore_errors=True)
    OUT.mkdir(parents=True)
    ASSETS.mkdir(parents=True, exist_ok=True)

    print("43-51 Chart rendering matrix")
    rows: list[tuple[RenderCase, Path, Path]] = []
    for case in CASES:
        if not case.source.exists():
            raise RuntimeError(f"Missing fixture for {case.key}: {case.source}")

        isolated = isolate_sheet(case)
        witan_png = ASSETS / f"case{case.number}_{case.key}_witan.png"
        libreoffice_png = ASSETS / f"case{case.number}_{case.key}_libreoffice.png"
        render_with_witan(isolated, case.range_address, witan_png)
        libreoffice_pdf_to_png(isolated, libreoffice_png)
        rows.append((case, witan_png, libreoffice_png))
        print(f"   {case.number} {case.key}: {case.family} / {case.variation}")
        print(f"      witan: {witan_png}")
        print(f"      LibreOffice: {libreoffice_png}")

    print()
    print(f"Summary: generated {len(rows)} Witan/LibreOffice chart render pairs")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
