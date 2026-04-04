# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

This repository automates publication-quality plot generation from HPLC/analytical chemistry experimental data. The user runs a single Python script against an Excel file; the script reads plot specifications from a `__plots__` sheet and outputs SVG files.

## Running the Script

### R 방식 (권장 — PDF, Illustrator 클리핑 마스크 없음)

```bash
# 최초 1회: R 패키지 설치
Rscript setup.R

# Excel 파일에서 PDF 그래프 생성
Rscript plot.R <experiment>.xlsx
```

Output PDF files are written to the same directory as the input Excel file, named after each `plot_id` in the `__plots__` sheet.

### Python 방식 (레거시 — SVG/PDF, 클리핑 마스크 문제 있음)

```bash
# First-time setup: create virtual environment and install dependencies
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

# Generate plots from an Excel file
.venv/bin/python3 plot.py <experiment>.xlsx
```

Output SVG files are written to the same directory as the input Excel file, named after each `plot_id` in the `__plots__` sheet.

## Excel File Convention

Every input Excel file must contain a `__plots__` sheet where each row defines one plot. Key columns:

| Column | Description |
|--------|-------------|
| `plot_id` | Output filename (no extension) |
| `plot_type` | `bar` / `line` / `scatter` / `dose_response` / `bar_line` |
| `data_sheet` | Sheet name containing the data |
| `x_col` | Column name for X axis |
| `y_cols` | Y value column(s); comma-separated for triplicates (auto mean/std) |
| `err_col` | Pre-calculated error column (used when `y_cols` is a single mean column) |
| `group_col` | Column for grouping/color separation |
| `x_label` / `y_label` | Axis labels |
| `x_scale` | `linear` (default) or `log` |
| `y_min` / `y_max` | Y-axis tick range (first and last tick); auto-computed if blank |

Data sheets use **wide format**: conditions as columns, measurements as rows. Triplicates are placed in adjacent columns (e.g., `WT_1`, `WT_2`, `WT_3`).

### `bar_line` 전용 추가 컬럼

하나의 chart에 bar + line이 공존하는 다축 복합 그래프. 각 series가 독립 Y축을 가짐.

| Column | Description | Example |
|--------|-------------|---------|
| `series_types` | 파이프(`\|`) 구분 series 타입 | `bar\|line\|line` |
| `series_y_cols` | 파이프 구분 Y 컬럼명 (콤마로 replicate 지정) | `r1,r2,r3\|OD_mean\|pH` |
| `series_err_cols` | 파이프 구분 error 컬럼 (replicate 사용 시 빈칸) | `\|OD_std\|` |
| `series_y_labels` | 파이프 구분 Y축 레이블 | `Indican (mM)\|OD600\|pH` |
| `series_names` | 파이프 구분 범례 이름 | `Indican\|OD600\|pH` |
| `series_y_mins` | 파이프 구분 각 series Y축 tick 하한 | `0\|0\|6.5` |
| `series_y_maxs` | 파이프 구분 각 series Y축 tick 상한 | `3\|1.5\|8` |

- Series 1 → 좌측 Y축, Series 2 → 우측 Y축, Series 3+ → 추가 우측 축 (65pt씩 오른쪽)
- 7개 이상 조건이 필요한 경우 `plot.py` 상단의 `PALETTE` 리스트에 색을 추가

## Architecture

`plot.py` is a single-file script organized as:

- **Config loading**: Parses `__plots__` sheet into a list of plot configs
- **Data loading**: Reads the specified sheet/columns; handles both raw triplicates and pre-calculated mean/std
- **Dispatch**: Routes each config to one of five plot functions (`plot_bar`, `plot_line`, `plot_scatter`, `plot_dose_response`, `plot_bar_line`)
- **Style**: A shared `apply_style()` function enforces Arial font, transparent background, spine removal, and the project color palette

## Output Style

- Format: SVG, transparent background
- Font: Arial (fixed)
- Color palette: `#5B9BD5`, `#ED7D31`, `#70AD47`, `#C45A77`, `#44ABAB`, `#8E6BB5`
- Bar charts include individual data points (dots) overlaid on the right side of each bar
- Axes: top and right spines removed

## Data Context

The Excel files contain HPLC experiments measuring Indican, Tryptophan (Trp), 3-IP, and related metabolites across conditions such as pH, temperature, time course, glucose concentration, and bacterial strains (WT vs. deletion mutants). All experiments use triplicate or fewer measurements.
