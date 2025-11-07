"""Analisis regresi satu variabel untuk sheet Onevar.

Membaca data dari file Excel tanpa dependensi eksternal dan menghasilkan
scatter plot (format SVG), korelasi Pearson, serta model regresi linear,
kuadratik, logaritmik, dan eksponensial.
"""
from __future__ import annotations

import argparse
import math
import statistics
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Sequence, Tuple
from xml.etree import ElementTree as ET


NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def load_sheet(path: Path, sheet_id: int = 1) -> List[List[str]]:
    """Load worksheet as a table of strings.

    Only supports basic shared strings and numeric cells, which is
    sufficient for the provided workbook.
    """

    with zipfile.ZipFile(path) as zf:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            shared_strings = [
                t.text for t in root.findall(".//main:si/main:t", namespaces=NS)
            ]
        sheet = ET.fromstring(zf.read(f"xl/worksheets/sheet{sheet_id}.xml"))

    data_rows: List[Dict[int, str]] = []
    for row in sheet.findall(".//main:sheetData/main:row", namespaces=NS):
        row_data: Dict[int, str] = {}
        for cell in row.findall("main:c", namespaces=NS):
            ref = cell.get("r")
            if ref is None:
                continue
            col_letters = "".join(filter(str.isalpha, ref))
            col_idx = 0
            for ch in col_letters:
                col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)
            value_elem = cell.find("main:v", namespaces=NS)
            value = value_elem.text if value_elem is not None else None
            if cell.get("t") == "s" and value is not None:
                value = shared_strings[int(value)]
            if value is not None:
                row_data[col_idx - 1] = value
        if row_data:
            data_rows.append(row_data)

    if not data_rows:
        return []

    num_cols = max(max(indices.keys()) for indices in data_rows) + 1
    table: List[List[str]] = []
    for row_data in data_rows:
        row = [row_data.get(col) for col in range(num_cols)]
        table.append(row)
    return table


def to_float_column(rows: Sequence[List[str]], idx: int) -> List[float]:
    return [float(row[idx]) for row in rows]


def pearson_corr(x: Sequence[float], y: Sequence[float]) -> float:
    if len(x) != len(y):
        raise ValueError("Input lengths must match")
    if len(x) < 2:
        raise ValueError("At least two data points required")

    mean_x = statistics.mean(x)
    mean_y = statistics.mean(y)
    num = sum((xi - mean_x) * (yi - mean_y) for xi, yi in zip(x, y))
    den_x = math.sqrt(sum((xi - mean_x) ** 2 for xi in x))
    den_y = math.sqrt(sum((yi - mean_y) ** 2 for yi in y))
    if den_x == 0 or den_y == 0:
        return float("nan")
    return num / (den_x * den_y)


@dataclass
class RegressionResult:
    name: str
    equation: str
    r_squared: float
    used_points: int
    parameters: Dict[str, float] | None = None


def r_squared(actual: Sequence[float], predicted: Sequence[float]) -> float:
    mean_y = statistics.mean(actual)
    ss_tot = sum((yi - mean_y) ** 2 for yi in actual)
    ss_res = sum((yi - pi) ** 2 for yi, pi in zip(actual, predicted))
    return 1 - ss_res / ss_tot if ss_tot != 0 else float("nan")


def linear_regression(x: Sequence[float], y: Sequence[float]) -> RegressionResult:
    n = len(x)
    sum_x = sum(x)
    sum_y = sum(y)
    sum_xx = sum(xi * xi for xi in x)
    sum_xy = sum(xi * yi for xi, yi in zip(x, y))

    denom = n * sum_xx - sum_x ** 2
    if denom == 0:
        raise ValueError("Cannot compute linear regression: singular matrix")
    slope = (n * sum_xy - sum_x * sum_y) / denom
    intercept = (sum_y - slope * sum_x) / n
    predicted = [intercept + slope * xi for xi in x]
    equation = f"y = {intercept:.6f} + {slope:.6f}x"
    params = {"intercept": intercept, "slope": slope}
    return RegressionResult("Linear", equation, r_squared(y, predicted), n, params)


def solve_3x3(a: List[List[float]], b: List[float]) -> Tuple[float, float, float]:
    """Solve a 3x3 linear system using Gaussian elimination."""

    # Augmented matrix
    m = [row[:] + [rhs] for row, rhs in zip(a, b)]
    for col in range(3):
        # Find pivot
        pivot_row = max(range(col, 3), key=lambda r: abs(m[r][col]))
        if abs(m[pivot_row][col]) < 1e-12:
            raise ValueError("Singular matrix in quadratic regression")
        if pivot_row != col:
            m[col], m[pivot_row] = m[pivot_row], m[col]
        # Normalize pivot row
        pivot = m[col][col]
        m[col] = [val / pivot for val in m[col]]
        # Eliminate other rows
        for r in range(3):
            if r == col:
                continue
            factor = m[r][col]
            m[r] = [rv - factor * pv for rv, pv in zip(m[r], m[col])]
    return (m[0][3], m[1][3], m[2][3])


def quadratic_regression(x: Sequence[float], y: Sequence[float]) -> RegressionResult:
    n = len(x)
    sum_x = sum(x)
    sum_x2 = sum(xi ** 2 for xi in x)
    sum_x3 = sum(xi ** 3 for xi in x)
    sum_x4 = sum(xi ** 4 for xi in x)
    sum_y = sum(y)
    sum_xy = sum(xi * yi for xi, yi in zip(x, y))
    sum_x2y = sum((xi ** 2) * yi for xi, yi in zip(x, y))

    a = [
        [n, sum_x, sum_x2],
        [sum_x, sum_x2, sum_x3],
        [sum_x2, sum_x3, sum_x4],
    ]
    b = [sum_y, sum_xy, sum_x2y]
    c0, c1, c2 = solve_3x3(a, b)
    predicted = [c0 + c1 * xi + c2 * xi ** 2 for xi in x]
    equation = f"y = {c0:.6f} + {c1:.6f}x + {c2:.6f}x^2"
    params = {"c0": c0, "c1": c1, "c2": c2}
    return RegressionResult("Quadratic", equation, r_squared(y, predicted), n, params)


def logarithmic_regression(x: Sequence[float], y: Sequence[float]) -> RegressionResult:
    filtered: List[Tuple[float, float]] = [
        (xi, yi) for xi, yi in zip(x, y) if xi > 0
    ]
    if len(filtered) < 2:
        raise ValueError("Not enough positive X values for logarithmic regression")
    log_x = [math.log(xi) for xi, _ in filtered]
    y_values = [yi for _, yi in filtered]
    base_result = linear_regression(log_x, y_values)
    assert base_result.parameters is not None
    intercept = base_result.parameters["intercept"]
    slope = base_result.parameters["slope"]
    predicted = [intercept + slope * math.log(xi) for xi, _ in filtered]
    equation = f"y = {intercept:.6f} + {slope:.6f} ln(x)"
    params = {"intercept": intercept, "slope": slope}
    return RegressionResult(
        "Logarithmic",
        equation,
        r_squared(y_values, predicted),
        len(filtered),
        params,
    )


def exponential_regression(x: Sequence[float], y: Sequence[float]) -> RegressionResult:
    filtered: List[Tuple[float, float]] = [
        (xi, yi) for xi, yi in zip(x, y) if yi > 0
    ]
    if len(filtered) < 2:
        raise ValueError("Not enough positive Y values for exponential regression")
    x_vals = [xi for xi, _ in filtered]
    log_y = [math.log(yi) for _, yi in filtered]
    base_result = linear_regression(x_vals, log_y)
    assert base_result.parameters is not None
    intercept = base_result.parameters["intercept"]
    slope = base_result.parameters["slope"]
    a = math.exp(intercept)
    b = slope
    predicted = [a * math.exp(b * xi) for xi in x_vals]
    equation = f"y = {a:.6f} * e^({b:.6f}x)"
    params = {"a": a, "b": b}
    return RegressionResult(
        "Exponential",
        equation,
        r_squared([yi for _, yi in filtered], predicted),
        len(filtered),
        params,
    )


def create_scatter_svg(
    x: Sequence[float],
    y: Sequence[float],
    title: str,
    x_label: str,
    y_label: str,
    path: Path,
    width: int = 640,
    height: int = 480,
) -> None:
    margin = 60
    inner_width = width - 2 * margin
    inner_height = height - 2 * margin
    min_x, max_x = min(x), max(x)
    min_y, max_y = min(y), max(y)
    # Avoid division by zero if constant values
    x_range = max_x - min_x if max_x != min_x else 1.0
    y_range = max_y - min_y if max_y != min_y else 1.0

    def scale_x(value: float) -> float:
        return margin + (value - min_x) / x_range * inner_width

    def scale_y(value: float) -> float:
        return height - margin - (value - min_y) / y_range * inner_height

    points = "\n".join(
        f'<circle cx="{scale_x(xi):.2f}" cy="{scale_y(yi):.2f}" r="3" fill="#1f77b4" />'
        for xi, yi in zip(x, y)
    )

    svg = f"""<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}'>
  <style>
    text {{ font-family: Arial, sans-serif; font-size: 14px; }}
    .title {{ font-size: 18px; font-weight: bold; }}
  </style>
  <rect x='0' y='0' width='{width}' height='{height}' fill='white'/>
  <line x1='{margin}' y1='{height - margin}' x2='{width - margin}' y2='{height - margin}' stroke='black'/>
  <line x1='{margin}' y1='{margin}' x2='{margin}' y2='{height - margin}' stroke='black'/>
  <text x='{width/2:.1f}' y='{margin/2:.1f}' text-anchor='middle' class='title'>{title}</text>
  <text x='{width/2:.1f}' y='{height - margin/3:.1f}' text-anchor='middle'>{x_label}</text>
  <text x='{margin/3:.1f}' y='{height/2:.1f}' text-anchor='middle' transform='rotate(-90 {margin/3:.1f},{height/2:.1f})'>{y_label}</text>
  {points}
</svg>"""
    path.write_text(svg, encoding="utf-8")


def analyze_variable(
    name: str,
    x: Sequence[float],
    y: Sequence[float],
) -> List[RegressionResult]:
    models: List[Callable[[Sequence[float], Sequence[float]], RegressionResult]] = [
        linear_regression,
        quadratic_regression,
        logarithmic_regression,
        exponential_regression,
    ]
    results: List[RegressionResult] = []
    for model in models:
        try:
            result = model(x, y)
            results.append(result)
        except ValueError as err:
            results.append(
                RegressionResult(
                    model.__name__.replace("_", " ").title(),
                    f"Model tidak dapat dihitung: {err}",
                    float("nan"),
                    0,
                )
            )
    return results


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Analisis regresi satu variabel menggunakan data dari file Excel "
            "dan menghasilkan laporan beserta scatter plot."
        )
    )
    parser.add_argument(
        "--data",
        type=Path,
        default=Path("Data Analisis Regresi Satu Variabel.xlsx"),
        help=(
            "Jalur ke file Excel sumber data. Secara bawaan menggunakan file "
            "di direktori skrip."
        ),
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("outputs"),
        help="Direktori tujuan untuk laporan dan plot yang dihasilkan.",
    )
    parser.add_argument(
        "--sheet-id",
        type=int,
        default=1,
        help="Nomor sheet (berbasis 1) yang akan diproses dari workbook.",
    )
    return parser


def main(args: Sequence[str] | None = None) -> None:
    parser = build_arg_parser()
    parsed = parser.parse_args(args)

    workbook = parsed.data
    if not workbook.exists():
        parser.error(f"File data tidak ditemukan: {workbook}")

    table = load_sheet(workbook, parsed.sheet_id)
    headers = table[0]
    rows = table[1:]

    indices = {name: headers.index(name) for name in ("X", "Y", "L", "R", "S")}
    x_values = to_float_column(rows, indices["X"])

    output_dir = parsed.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)
    plot_dir = output_dir / "plots"
    plot_dir.mkdir(exist_ok=True)

    report_lines: List[str] = []
    report_lines.append("Analisis regresi sheet Onevar")
    report_lines.append("")

    for target in ("Y", "L", "R", "S"):
        y_values = to_float_column(rows, indices[target])
        corr = pearson_corr(x_values, y_values)
        report_lines.append(f"Variabel dependen: {target}")
        report_lines.append(f"  Korelasi Pearson (X vs {target}): {corr:.6f}")

        plot_path = plot_dir / f"scatter_X_vs_{target}.svg"
        create_scatter_svg(
            x_values,
            y_values,
            f"Scatter plot X vs {target}",
            "X",
            target,
            plot_path,
        )
        report_lines.append(f"  Scatter plot: {plot_path}")

        results = analyze_variable(target, x_values, y_values)
        for result in results:
            usable = f" (n={result.used_points})" if result.used_points else ""
            r2_display = "nan" if math.isnan(result.r_squared) else f"{result.r_squared:.6f}"
            report_lines.append(f"    {result.name}{usable}: {result.equation} | R^2={r2_display}")
        best = max(
            (r for r in results if not math.isnan(r.r_squared)),
            key=lambda r: r.r_squared,
            default=None,
        )
        if best is None:
            report_lines.append("    Tidak ada model yang valid.")
        else:
            report_lines.append(
                f"    Model terbaik berdasarkan R^2: {best.name} (R^2={best.r_squared:.6f})"
            )
        report_lines.append("")

    report_path = output_dir / "onevar_regression_summary.txt"
    report_path.write_text("\n".join(report_lines), encoding="utf-8")


if __name__ == "__main__":
    main()
