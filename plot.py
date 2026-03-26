#!/usr/bin/env python3
"""
plot.py — Excel 파일에서 publication-quality SVG 그래프를 자동 생성한다.

사용법:
    .venv/bin/python3 plot.py <experiment>.xlsx

입력 Excel 파일에 '__plots__' 시트가 반드시 있어야 하며,
각 행이 SVG 1개를 정의한다. 출력 SVG는 Excel 파일과 같은 디렉터리에 저장된다.

──────── __plots__ 시트 컬럼 규약 ────────
공통 (모든 plot_type):
  plot_id        출력 파일명 (확장자 제외)
  plot_type      bar / line / scatter / dose_response / bar_line
  data_sheet     데이터가 있는 시트명
  data_start_row 헤더 행 번호(1-indexed). 생략 시 1행
  x_col          X축 컬럼명
  y_cols         Y값 컬럼명. 콤마 구분 복수 지정 시 triplicate → 자동 mean/std 계산
  err_col        사전 계산된 오차 컬럼 (y_cols가 단일 mean 컬럼일 때)
  group_col      그룹/색상 구분 컬럼 (grouped bar, multi-line에 사용)
  x_label        X축 레이블
  y_label        Y축 레이블
  title          그래프 제목 (선택)
  x_scale        log 입력 시 로그 스케일 (기본: linear)

bar_line 전용 (series_* 컬럼, 파이프 | 로 series 구분):
  series_types     각 series 타입  예) bar|line|line
  series_y_cols    각 series Y 컬럼 (콤마로 replicate 지정)  예) r1,r2,r3|OD_mean|pH
  series_err_cols  각 series 오차 컬럼 (replicate 사용 시 빈칸)  예) |OD_std|
  series_y_labels  각 series Y축 레이블  예) Indican (mM)|OD600|pH
  series_names     각 series 범례 이름  예) Indican|OD600|pH
  * Series 1 → 좌측 Y축, Series 2 → 우측 Y축, Series 3+ → 추가 우측 축 (65pt씩)
"""

import re
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from scipy import stats

warnings.filterwarnings("ignore")
matplotlib.use("Agg")
matplotlib.rcParams['svg.fonttype'] = 'path'  # 텍스트를 패스로 변환 (Illustrator 호환)
matplotlib.rcParams['pdf.fonttype'] = 42       # TrueType 임베드 (Illustrator 호환)

# ── Style constants ────────────────────────────────────────────────────────────

PALETTE = [
    "#3E9DAA",  # steel teal   (183°)
    "#C4607A",  # dusty rose   (344°)
    "#2B6840",  # forest green (143°)
    "#BF9A2A",  # ochre        ( 40°)
    "#A83225",  # brick red    (  6°)
    "#4E78A0",  # slate blue   (210°)
    "#879A28",  # moss olive   ( 71°)
    "#6A3E5A",  # dark plum    (310°)
]

FS = {"title": 12, "label": 11, "tick": 10, "legend": 9, "annot": 9}
FONT = "Arial"


# ── Config loading ─────────────────────────────────────────────────────────────

def load_configs(xlsx_path: Path) -> list:
    xl = pd.ExcelFile(xlsx_path)
    if "__plots__" not in xl.sheet_names:
        raise ValueError("No '__plots__' sheet found. Add a '__plots__' sheet with plot specifications.")

    df = xl.parse("__plots__", dtype=str).fillna("")
    configs = []
    for _, row in df.iterrows():
        cfg = {k.strip(): v.strip() for k, v in row.items()}
        if cfg.get("plot_id"):
            configs.append(cfg)
    return configs


# ── Data loading ───────────────────────────────────────────────────────────────

def load_sheet(xlsx_path: Path, cfg: dict) -> pd.DataFrame:
    sheet = cfg.get("data_sheet") or "Sheet1"
    skip_raw = cfg.get("data_start_row", "").strip()
    # data_start_row is 1-indexed row number of the HEADER row
    skiprows = (int(skip_raw) - 1) if skip_raw else None

    xl = pd.ExcelFile(xlsx_path)
    df = xl.parse(sheet, skiprows=skiprows, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all")


def extract_series(df: pd.DataFrame, cfg: dict):
    """Return (x_vals, y_mean, y_err, y_reps_list, group_vals).

    y_reps_list : list of lists (one per replicate column), or None
    y_err       : list of floats (std), or list of None
    group_vals  : list of str labels per row, or None
    """
    x_col     = cfg.get("x_col", "")
    y_cols_s  = cfg.get("y_cols", "")
    err_col   = cfg.get("err_col", "")
    group_col = cfg.get("group_col", "")

    y_cols = [c.strip() for c in y_cols_s.split(",") if c.strip()]

    x_vals = df[x_col].tolist() if x_col in df.columns else list(range(len(df)))

    group_vals = (
        df[group_col].astype(str).tolist()
        if group_col and group_col in df.columns
        else None
    )

    if len(y_cols) > 1:
        # Raw replicates → compute mean / std automatically
        rep_df = df[y_cols].apply(pd.to_numeric, errors="coerce")
        y_mean = rep_df.mean(axis=1).tolist()
        y_err  = rep_df.std(axis=1, ddof=1).tolist()
        y_reps = [rep_df.iloc[:, i].tolist() for i in range(rep_df.shape[1])]
    else:
        y_col  = y_cols[0] if y_cols else None
        y_mean = pd.to_numeric(df[y_col], errors="coerce").tolist() if y_col else []
        y_reps = None
        if err_col and err_col in df.columns:
            y_err = pd.to_numeric(df[err_col], errors="coerce").tolist()
        else:
            y_err = [None] * len(y_mean)

    return x_vals, y_mean, y_err, y_reps, group_vals


# ── Shared style helpers ───────────────────────────────────────────────────────

def _apply_style(ax, cfg: dict):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    x_label = cfg.get("x_label", "")
    y_label = cfg.get("y_label", "")
    title   = cfg.get("title", "")

    if x_label:
        ax.set_xlabel(x_label, fontsize=FS["label"], fontfamily=FONT)
    if y_label:
        ax.set_ylabel(y_label, fontsize=FS["label"], fontfamily=FONT)
    if title:
        ax.set_title(title, fontsize=FS["title"], fontfamily=FONT)

    ax.tick_params(labelsize=FS["tick"])
    for lbl in ax.get_xticklabels() + ax.get_yticklabels():
        lbl.set_fontfamily(FONT)

    if cfg.get("x_scale") == "log":
        ax.set_xscale("log")


def _save(fig, out_dir: Path, plot_id: str, fmt: str = "svg"):
    out = out_dir / f"{plot_id}.{fmt}"
    if fmt == "pdf":
        fig.savefig(out, format="pdf", transparent=True, bbox_inches="tight")
    else:
        fig.savefig(out, format="svg", transparent=True, bbox_inches="tight", dpi=300)
        _strip_clippath(out)
    plt.close(fig)
    print(f"    Saved → {out.name}")


def _strip_clippath(svg_path: Path):
    """Illustrator 'Tiny 변환 시 클리핑 손실' 경고 원인인 clipPath 요소를 제거."""
    txt = svg_path.read_text(encoding="utf-8")
    txt = re.sub(r'<clipPath\b[^>]*>.*?</clipPath>', '', txt, flags=re.DOTALL)
    txt = re.sub(r'\s+clip-path="[^"]*"', '', txt)
    svg_path.write_text(txt, encoding="utf-8")


def _has_errors(err_list) -> bool:
    return any(
        e is not None and not (isinstance(e, float) and np.isnan(e))
        for e in err_list
    )


def _err_kw() -> dict:
    return {"linewidth": 1.2, "capthick": 1.2, "capsize": 4}


# ── Bar chart ──────────────────────────────────────────────────────────────────

def plot_bar(xlsx_path: Path, cfg: dict, out_dir: Path, fmt: str = "svg"):
    df = load_sheet(xlsx_path, cfg)
    x_vals, y_mean, y_err, y_reps, group_vals = extract_series(df, cfg)

    n_x   = len(list(dict.fromkeys(x_vals)))
    n_g   = len(list(dict.fromkeys(group_vals))) if group_vals else 1
    fig_w = max(3.5, n_x * 0.5 + n_g * 0.4 + 1.0)
    fig, ax = plt.subplots(figsize=(fig_w, 4))

    if group_vals is not None:
        _grouped_bar(ax, x_vals, y_mean, y_err, y_reps, group_vals)
    else:
        _simple_bar(ax, x_vals, y_mean, y_err, y_reps)

    _apply_style(ax, cfg)
    fig.tight_layout()
    _save(fig, out_dir, cfg["plot_id"], fmt)


def _simple_bar(ax, x_vals, y_mean, y_err, y_reps):
    n      = len(y_mean)
    x_pos  = list(range(n))
    color  = PALETTE[0]
    width  = 0.6

    err_arg = y_err if _has_errors(y_err) else None
    ax.bar(x_pos, y_mean, width=width, color=color, alpha=0.85,
           yerr=err_arg, error_kw=_err_kw(), zorder=2)

    if y_reps:
        dot_offset = width * 0.4
        for rep_vals in y_reps:
            dot_xs = [xi + dot_offset for xi in x_pos]
            _clean_vals = [v if (v is not None and not (isinstance(v, float) and np.isnan(v))) else None
                           for v in rep_vals]
            ax.scatter(
                [dx for dx, v in zip(dot_xs, _clean_vals) if v is not None],
                [v for v in _clean_vals if v is not None],
                color="#333333", s=18, zorder=4, alpha=0.9,
                linewidths=0.5, edgecolors="white",
            )

    ax.set_xticks(x_pos)
    ax.set_xticklabels([str(v) for v in x_vals], fontfamily=FONT, fontsize=FS["tick"])


def _grouped_bar(ax, x_vals, y_mean, y_err, y_reps, group_vals):
    unique_x = list(dict.fromkeys(x_vals))
    unique_g = list(dict.fromkeys(group_vals))
    n_x      = len(unique_x)
    n_g      = len(unique_g)
    width    = 0.7 / n_g
    x_idx    = {v: i for i, v in enumerate(unique_x)}

    for gi, grp in enumerate(unique_g):
        color  = PALETTE[gi % len(PALETTE)]
        mask   = [g == grp for g in group_vals]
        gx_idx = [x_idx[x] for x, m in zip(x_vals, mask) if m]
        gy     = [y for y, m in zip(y_mean, mask) if m]
        ge     = [e for e, m in zip(y_err,  mask) if m]
        offsets = [xi + (gi - n_g / 2 + 0.5) * width for xi in gx_idx]

        err_arg = ge if _has_errors(ge) else None
        ax.bar(offsets, gy, width=width, color=color, alpha=0.85,
               yerr=err_arg, error_kw=_err_kw(), label=grp, zorder=2)

        if y_reps:
            dot_offset = width * 0.35
            for rep_vals in y_reps:
                g_rep = [v for v, m in zip(rep_vals, mask) if m]
                dot_xs = [o + dot_offset for o in offsets]
                ax.scatter(
                    dot_xs, g_rep,
                    color=color, s=18, zorder=4, alpha=0.9,
                    linewidths=0.5, edgecolors="white",
                )

    ax.set_xticks(range(n_x))
    ax.set_xticklabels(unique_x, fontfamily=FONT, fontsize=FS["tick"])
    ax.legend(fontsize=FS["legend"], prop={"family": FONT})


# ── Line chart ─────────────────────────────────────────────────────────────────

def plot_line(xlsx_path: Path, cfg: dict, out_dir: Path, fmt: str = "svg"):
    df = load_sheet(xlsx_path, cfg)
    x_vals, y_mean, y_err, y_reps, group_vals = extract_series(df, cfg)

    x_num = pd.to_numeric(pd.Series(x_vals), errors="coerce")

    fig, ax = plt.subplots(figsize=(5, 4))

    if group_vals is not None:
        unique_g = list(dict.fromkeys(group_vals))
        for gi, grp in enumerate(unique_g):
            color = PALETTE[gi % len(PALETTE)]
            mask  = [g == grp for g in group_vals]
            gx    = [x for x, m in zip(x_num, mask) if m]
            gy    = [y for y, m in zip(y_mean, mask) if m]
            ge    = [e for e, m in zip(y_err,  mask) if m]

            ax.plot(gx, gy, "o-", color=color, label=grp, linewidth=1.5, markersize=5)
            if _has_errors(ge):
                ax.errorbar(gx, gy, yerr=ge, fmt="none", color=color, **_err_kw())

        ax.legend(fontsize=FS["legend"], prop={"family": FONT})
    else:
        color = PALETTE[0]
        ax.plot(x_num, y_mean, "o-", color=color, linewidth=1.5, markersize=5)
        if _has_errors(y_err):
            ax.errorbar(x_num, y_mean, yerr=y_err, fmt="none", color=color, **_err_kw())

    _apply_style(ax, cfg)
    fig.tight_layout()
    _save(fig, out_dir, cfg["plot_id"], fmt)


# ── Scatter + Linear regression ────────────────────────────────────────────────

def plot_scatter(xlsx_path: Path, cfg: dict, out_dir: Path, fmt: str = "svg"):
    df = load_sheet(xlsx_path, cfg)

    x_col    = cfg.get("x_col", "")
    y_cols_s = cfg.get("y_cols", "")
    y_cols   = [c.strip() for c in y_cols_s.split(",") if c.strip()]
    y_col    = y_cols[0] if y_cols else None

    x_raw = pd.to_numeric(df[x_col], errors="coerce") if x_col in df.columns else pd.Series(dtype=float)
    y_raw = pd.to_numeric(df[y_col], errors="coerce") if y_col and y_col in df.columns else pd.Series(dtype=float)

    valid  = x_raw.notna() & y_raw.notna()
    x_data = x_raw[valid].values
    y_data = y_raw[valid].values

    fig, ax = plt.subplots(figsize=(5, 4))
    ax.scatter(x_data, y_data, color=PALETTE[0], s=45, alpha=0.9,
               linewidths=0.5, edgecolors="white", zorder=3)

    if len(x_data) >= 2:
        slope, intercept, r, _, _ = stats.linregress(x_data, y_data)
        x_line = np.linspace(x_data.min(), x_data.max(), 200)
        ax.plot(x_line, slope * x_line + intercept,
                "-", color=PALETTE[1], linewidth=1.5, alpha=0.9, zorder=2)

        sign   = "+" if intercept >= 0 else "-"
        eq_str = (
            f"$y = {slope:.4g}x {sign} {abs(intercept):.4g}$\n"
            f"$R^2 = {r**2:.4f}$"
        )
        ax.annotate(eq_str, xy=(0.05, 0.95), xycoords="axes fraction",
                    fontsize=FS["annot"], fontfamily=FONT,
                    verticalalignment="top", linespacing=1.6)

    _apply_style(ax, cfg)
    fig.tight_layout()
    _save(fig, out_dir, cfg["plot_id"], fmt)


# ── Dose-response (log-scale line) ────────────────────────────────────────────

def plot_dose_response(xlsx_path: Path, cfg: dict, out_dir: Path, fmt: str = "svg"):
    cfg = dict(cfg)
    cfg["x_scale"] = "log"
    # Filter x=0 rows before plotting (log scale cannot handle 0)
    df = load_sheet(xlsx_path, cfg)
    x_col = cfg.get("x_col", "")
    if x_col in df.columns:
        df = df[pd.to_numeric(df[x_col], errors="coerce") > 0]
    # Delegate to line chart logic
    _line_from_df(df, xlsx_path, cfg, out_dir, fmt)


def _line_from_df(df: pd.DataFrame, xlsx_path: Path, cfg: dict, out_dir: Path, fmt: str = "svg"):
    """Like plot_line but accepts a pre-filtered DataFrame."""
    x_vals, y_mean, y_err, y_reps, group_vals = extract_series(df, cfg)
    x_num = pd.to_numeric(pd.Series(x_vals), errors="coerce")

    fig, ax = plt.subplots(figsize=(5, 4))

    if group_vals is not None:
        unique_g = list(dict.fromkeys(group_vals))
        for gi, grp in enumerate(unique_g):
            color = PALETTE[gi % len(PALETTE)]
            mask  = [g == grp for g in group_vals]
            gx    = [x for x, m in zip(x_num, mask) if m]
            gy    = [y for y, m in zip(y_mean, mask) if m]
            ge    = [e for e, m in zip(y_err,  mask) if m]
            ax.plot(gx, gy, "o-", color=color, label=grp, linewidth=1.5, markersize=5)
            if _has_errors(ge):
                ax.errorbar(gx, gy, yerr=ge, fmt="none", color=color, **_err_kw())
        ax.legend(fontsize=FS["legend"], prop={"family": FONT})
    else:
        color = PALETTE[0]
        ax.plot(x_num, y_mean, "o-", color=color, linewidth=1.5, markersize=5)
        if _has_errors(y_err):
            ax.errorbar(x_num, y_mean, yerr=y_err, fmt="none", color=color, **_err_kw())

    _apply_style(ax, cfg)
    fig.tight_layout()
    _save(fig, out_dir, cfg["plot_id"], fmt)


# ── Bar-Line 복합 플롯 ──────────────────────────────────────────────────────────

def _parse_bar_line_series(cfg: dict) -> list:
    """pipe 구분된 series_* 컬럼을 series dict 리스트로 파싱."""
    def split_pipe(s):
        return [x.strip() for x in s.split("|")]

    types      = split_pipe(cfg.get("series_types", ""))
    y_cols_all = split_pipe(cfg.get("series_y_cols", ""))
    err_cols   = split_pipe(cfg.get("series_err_cols", ""))
    y_labels   = split_pipe(cfg.get("series_y_labels", ""))
    names      = split_pipe(cfg.get("series_names", ""))

    n = len(types)

    def pad(lst):
        return lst + [""] * max(0, n - len(lst))

    y_cols_all, err_cols, y_labels, names = (
        pad(y_cols_all), pad(err_cols), pad(y_labels), pad(names)
    )

    series = []
    for i in range(n):
        y_cols_i = [c.strip() for c in y_cols_all[i].split(",") if c.strip()]
        series.append({
            "type":    types[i].lower(),
            "y_cols":  y_cols_i,
            "err_col": err_cols[i],
            "y_label": y_labels[i],
            "name":    names[i],
        })
    return series


def _series_bar_on_ax(ax, x_pos, y_mean, y_err, y_reps, color, name):
    """bar_line용 bar series. 좌우 여백을 위해 고정 width 사용."""
    width   = 0.4
    err_arg = y_err if _has_errors(y_err) else None
    bars    = ax.bar(x_pos, y_mean, width=width, color=color, alpha=0.85,
                     yerr=err_arg, error_kw=_err_kw(), label=name, zorder=2)

    if y_reps:
        dot_offset = width * 0.4
        for rep_vals in y_reps:
            valid = [(xi + dot_offset, v)
                     for xi, v in zip(x_pos, rep_vals)
                     if v is not None and not (isinstance(v, float) and np.isnan(v))]
            if valid:
                dxs, dvs = zip(*valid)
                ax.scatter(dxs, dvs, color=color, s=18, zorder=4, alpha=0.9,
                           linewidths=0.5, edgecolors="white")
    return bars


def _series_line_on_ax(ax, x_pos, y_mean, y_err, color, name):
    """bar_line용 line series."""
    x_num = pd.to_numeric(pd.Series(x_pos), errors="coerce")
    line, = ax.plot(x_num, y_mean, "o-", color=color, label=name,
                    linewidth=1.5, markersize=5)
    if _has_errors(y_err):
        ax.errorbar(x_num, y_mean, yerr=y_err, fmt="none", color=color, **_err_kw())
    return line


def plot_bar_line(xlsx_path: Path, cfg: dict, out_dir: Path, fmt: str = "svg"):
    """하나의 chart 안에 bar + line이 공존하는 다축 복합 그래프.

    각 series는 독립적인 Y축을 가짐:
      Series 1 → 좌측 Y축
      Series 2 → 우측 Y축 (첫 번째 right axis)
      Series 3+ → 추가 right axis (65pt씩 바깥으로 배치)
    """
    df = load_sheet(xlsx_path, cfg)
    x_col  = cfg.get("x_col", "")
    x_vals = df[x_col].tolist() if x_col in df.columns else list(range(len(df)))

    # 항상 정수 위치 사용: bar width(0.4)가 x축 스케일과 무관하게 일정하게 표시됨
    x_pos = list(range(len(x_vals)))

    series_list = _parse_bar_line_series(cfg)
    n_series    = len(series_list)

    # 우측 axis가 늘어날수록 figure를 넓힘
    n_right  = max(0, n_series - 1)
    n_x      = len(x_vals)
    fig_w    = max(5.0, n_x * 0.6 + 1.5 + n_right * 0.9)
    fig, ax0 = plt.subplots(figsize=(fig_w, 4))

    axes = [ax0]
    for i in range(1, n_series):
        axi = ax0.twinx()
        if i > 1:
            axi.spines["right"].set_position(("outward", (i - 1) * 65))
        axes.append(axi)

    handles, legend_labels = [], []

    for i, (series, ax) in enumerate(zip(series_list, axes)):
        color  = PALETTE[i % len(PALETTE)]
        y_cols = series["y_cols"]
        err_col = series["err_col"]

        # Y 데이터 추출
        if len(y_cols) > 1:
            rep_df = df[y_cols].apply(pd.to_numeric, errors="coerce")
            y_mean = rep_df.mean(axis=1).tolist()
            y_err  = rep_df.std(axis=1, ddof=1).tolist()
            y_reps = [rep_df.iloc[:, j].tolist() for j in range(rep_df.shape[1])]
        elif len(y_cols) == 1:
            y_mean = pd.to_numeric(df[y_cols[0]], errors="coerce").tolist()
            y_reps = None
            if err_col and err_col in df.columns:
                y_err = pd.to_numeric(df[err_col], errors="coerce").tolist()
            else:
                y_err = [None] * len(y_mean)
        else:
            continue

        s_type = series["type"]
        if s_type == "bar":
            h = _series_bar_on_ax(ax, x_pos, y_mean, y_err, y_reps, color, series["name"])
        else:
            h = _series_line_on_ax(ax, x_pos, y_mean, y_err, color, series["name"])

        if h is not None and series["name"]:
            handles.append(h)
            legend_labels.append(series["name"])

        ax.set_ylabel(series["y_label"], fontsize=FS["label"], fontfamily=FONT)
        ax.tick_params(axis="y", labelsize=FS["tick"])

    # 모든 Y축 tick 위치 정렬: linspace로 동일 비율 위치 → 시각적 정렬 보장
    for ax in axes:
        ymin, ymax = ax.get_ylim()
        ax.set_yticks(np.linspace(ymin, ymax, 5))
        for lbl in ax.get_yticklabels():
            lbl.set_fontfamily(FONT)

    # 공통 X축 스타일 (twinx 축 포함 모든 축의 top spine 제거)
    for ax in axes:
        ax.spines["top"].set_visible(False)
    ax0.tick_params(axis="x", labelsize=FS["tick"])
    for lbl in ax0.get_xticklabels():
        lbl.set_fontfamily(FONT)

    ax0.set_xticks(x_pos)
    ax0.set_xticklabels([str(v) for v in x_vals],
                        fontfamily=FONT, fontsize=FS["tick"])

    x_label = cfg.get("x_label", "")
    if x_label:
        ax0.set_xlabel(x_label, fontsize=FS["label"], fontfamily=FONT)

    title = cfg.get("title", "")
    if title:
        ax0.set_title(title, fontsize=FS["title"], fontfamily=FONT)

    if handles:
        ax0.legend(handles, legend_labels,
                   fontsize=FS["legend"], prop={"family": FONT}, loc="upper left")

    fig.tight_layout()
    _save(fig, out_dir, cfg["plot_id"], fmt)


# ── Dispatch ───────────────────────────────────────────────────────────────────

DISPATCH = {
    "bar":           plot_bar,
    "line":          plot_line,
    "scatter":       plot_scatter,
    "dose_response": plot_dose_response,
    "bar_line":      plot_bar_line,
}


# ── Entry point ────────────────────────────────────────────────────────────────

def main(xlsx_path_str: str, fmt: str = "pdf"):
    xlsx_path = Path(xlsx_path_str).resolve()
    if not xlsx_path.exists():
        print(f"Error: file not found — {xlsx_path}")
        sys.exit(1)

    out_dir = xlsx_path.parent
    print(f"\nReading: {xlsx_path.name}")

    configs = load_configs(xlsx_path)
    print(f"Found {len(configs)} plot(s) in '__plots__' sheet.\n")

    ok = err = 0
    for cfg in configs:
        plot_id   = cfg.get("plot_id", "?")
        plot_type = cfg.get("plot_type", "").lower()
        print(f"  [{plot_type}] {plot_id}")

        fn = DISPATCH.get(plot_type)
        if fn is None:
            print(f"    Warning: unknown plot_type '{plot_type}', skipping.")
            continue

        try:
            fn(xlsx_path, cfg, out_dir, fmt)
            ok += 1
        except Exception as e:
            print(f"    Error: {e}")
            err += 1

    print(f"\nDone — {ok} succeeded, {err} failed.")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Excel 파일에서 publication-quality 그래프를 생성한다.")
    parser.add_argument("xlsx", help="입력 Excel 파일 경로")
    parser.add_argument(
        "--format", dest="fmt", choices=["svg", "pdf"], default="pdf",
        help="출력 형식 (기본값: pdf)"
    )
    args = parser.parse_args()
    main(args.xlsx, fmt=args.fmt)
