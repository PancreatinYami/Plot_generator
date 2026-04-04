#!/usr/bin/env Rscript
# plot.R — Excel 파일에서 publication-quality PDF 그래프를 자동 생성한다.
#
# 사용법:
#   Rscript plot.R <experiment>.xlsx
#
# 입력 Excel 파일에 '__plots__' 시트가 반드시 있어야 하며,
# 각 행이 PDF 1개를 정의한다. 출력 PDF는 Excel 파일과 같은 디렉터리에 저장된다.
#
# ──────── __plots__ 시트 컬럼 규약 ────────
# 공통 (모든 plot_type):
#   plot_id        출력 파일명 (확장자 제외)
#   plot_type      bar / line / scatter / dose_response / bar_line
#   data_sheet     데이터가 있는 시트명
#   data_start_row 헤더 행 번호(1-indexed). 생략 시 1행
#   x_col          X축 컬럼명
#   y_cols         Y값 컬럼명. 콤마 구분 복수 지정 시 triplicate → 자동 mean/sd 계산
#   err_col        사전 계산된 오차 컬럼 (y_cols가 단일 mean 컬럼일 때)
#   group_col      그룹/색상 구분 컬럼 (grouped bar, multi-line에 사용)
#   x_label        X축 레이블
#   y_label        Y축 레이블
#   title          그래프 제목 (선택)
#   x_scale        log 입력 시 로그 스케일 (기본: linear)
#   y_min          Y축 최솟값 (빈칸이면 자동)
#   y_max          Y축 최댓값 (빈칸이면 자동)
#
# bar_line 전용 (series_* 컬럼, 파이프 | 로 series 구분):
#   series_types     각 series 타입  예) bar|line|line
#   series_y_cols    각 series Y 컬럼 (콤마로 replicate 지정)  예) r1,r2,r3|OD_mean|pH
#   series_err_cols  각 series 오차 컬럼 (replicate 사용 시 빈칸)  예) |OD_std|
#   series_y_labels  각 series Y축 레이블  예) Indican (mM)|OD600|pH
#   series_names     각 series 범례 이름  예) Indican|OD600|pH
#   series_y_mins    각 series Y 최솟값 (빈칸이면 자동)  예) 0|0|6.5
#   series_y_maxs    각 series Y 최댓값 (빈칸이면 자동)  예) 3|1.5|8
#   * Series 1 → 좌측 Y축, Series 2 → 우측 Y축, Series 3+ → 추가 우측 축

suppressPackageStartupMessages({
  library(readxl)
  library(ggplot2)
  library(scales)
  library(cowplot)
})

# ── Style constants ────────────────────────────────────────────────────────────

PALETTE <- c(
  "#3E9DAA",  # steel teal
  "#C4607A",  # dusty rose
  "#2B6840",  # forest green
  "#BF9A2A",  # ochre
  "#A83225",  # brick red
  "#4E78A0",  # slate blue
  "#879A28",  # moss olive
  "#6A3E5A"   # dark plum
)

FS   <- list(title=12, label=11, tick=10, legend=9, annot=9)

# Detect best available PDF font: prefer Arial → Helvetica → sans
.pdf_fonts <- local({
  tmp <- tempfile(fileext=".pdf")
  pdf(tmp, width=1, height=1)
  fns <- names(pdfFonts())
  dev.off()
  unlink(tmp)
  fns
})
FONT <- if ("Arial" %in% .pdf_fonts) "Arial" else
        if ("Helvetica" %in% .pdf_fonts) "Helvetica" else "sans"
rm(.pdf_fonts)

# ggplot2 annotate("text", size=) uses mm, not pt.
# Conversion: size_mm = pt / 2.845276
PT <- 2.845276

# NULL-coalescing operator
`%||%` <- function(a, b) if (!is.null(a)) a else b

# Top tick appears at this fraction of the panel height (same for all axes → aligned).
# E.g. 0.85 → top tick at 85 % of panel height, 15 % whitespace above.
TOP_TICK_FRAC <- 0.85

# Compute tick breaks and panel limits for one Y-axis.
#
#   data_lo / data_hi : actual data range (used when user hasn't specified limits)
#   user_lo / user_hi : exact tick endpoints set by the user (NULL = auto)
#   n                 : hint for number of ticks (passed to pretty())
#
# Returns list(breaks, lims):
#   breaks — clean, equally-spaced tick values
#   lims   — c(panel_lo, panel_hi); panel_hi is above max(breaks) so there is
#            whitespace at the top; panel_hi / panel_lo fraction is TOP_TICK_FRAC
#            for every axis → all top ticks at the same visual height.
make_axis_layout <- function(data_lo, data_hi, user_lo = NULL, user_hi = NULL, n = 5L) {
  tick_lo <- user_lo %||% data_lo
  tick_hi <- user_hi %||% data_hi
  if (isTRUE(all.equal(tick_lo, tick_hi))) tick_hi <- tick_lo + 1

  brks <- pretty(c(tick_lo, tick_hi), n = n)

  if (!is.null(user_lo) || !is.null(user_hi)) {
    # User-specified: force exact endpoints; pretty() fills clean intermediate values.
    lo_bound <- user_lo %||% min(brks)
    hi_bound <- user_hi %||% max(brks)
    inner    <- brks[brks > lo_bound & brks < hi_bound]
    brks     <- sort(unique(c(lo_bound, inner, hi_bound)))
  } else {
    # Auto: let pretty() decide the lower bound naturally (avoids forcing ugly
    # raw minimums like 0.045 or 0.09 as tick labels).
    brks <- brks[brks <= data_hi]
    if (length(brks) == 0L) brks <- c(tick_lo, tick_hi)
  }

  panel_lo <- min(brks)
  top_tick <- max(brks)
  # panel_hi is chosen so that top_tick sits at exactly TOP_TICK_FRAC of the panel.
  # Because the fraction is constant across all axes, all top ticks align visually.
  panel_hi <- panel_lo + (top_tick - panel_lo) / TOP_TICK_FRAC
  list(breaks = brks, lims = c(panel_lo, panel_hi))
}

# Parse an optional numeric limit from Excel cfg (empty string → NULL).
parse_ylim <- function(val) {
  if (is.null(val) || is.na(val) || trimws(val) == "") return(NULL)
  v <- suppressWarnings(as.numeric(val))
  if (is.na(v)) NULL else v
}

theme_pub <- function() {
  theme_classic(base_size=FS$tick, base_family=FONT) +
  theme(
    plot.title        = element_text(size=FS$title,  family=FONT, hjust=0.5),
    axis.title.x      = element_text(size=FS$label,  family=FONT, margin=margin(t=8)),
    axis.title.y      = element_text(size=FS$label,  family=FONT, margin=margin(r=8)),
    axis.text         = element_text(size=FS$tick,   family=FONT),
    legend.text       = element_text(size=FS$legend, family=FONT),
    legend.title      = element_blank(),
    legend.background = element_blank(),
    legend.key        = element_blank()
  )
}

# ── Config loading ─────────────────────────────────────────────────────────────

load_configs <- function(xlsx_path) {
  sheets <- excel_sheets(xlsx_path)
  if (!"__plots__" %in% sheets) {
    stop("No '__plots__' sheet found. Add a '__plots__' sheet with plot specifications.")
  }
  df <- read_excel(xlsx_path, sheet="__plots__", col_types="text")
  df[is.na(df)] <- ""
  colnames(df)  <- trimws(colnames(df))
  df <- as.data.frame(lapply(df, trimws), stringsAsFactors=FALSE)
  df <- df[df$plot_id != "", , drop=FALSE]
  lapply(seq_len(nrow(df)), function(i) as.list(df[i, , drop=FALSE]))
}

# ── Data loading ───────────────────────────────────────────────────────────────

load_sheet <- function(xlsx_path, cfg) {
  sheet    <- if (cfg$data_sheet != "") cfg$data_sheet else "Sheet1"
  skip_raw <- cfg$data_start_row
  skip     <- if (skip_raw != "" && !is.na(suppressWarnings(as.integer(skip_raw))))
    as.integer(skip_raw) - 1L else 0L

  df <- read_excel(xlsx_path, sheet=sheet, skip=skip, col_names=TRUE)
  colnames(df) <- trimws(colnames(df))
  df <- df[rowSums(!is.na(df)) > 0, , drop=FALSE]
  as.data.frame(df, stringsAsFactors=FALSE)
}

extract_series <- function(df, cfg) {
  x_col     <- cfg$x_col
  y_cols_s  <- cfg$y_cols
  err_col   <- cfg$err_col
  group_col <- cfg$group_col

  y_cols <- trimws(strsplit(y_cols_s, ",")[[1]])
  y_cols <- y_cols[y_cols != ""]

  x_vals <- if (x_col != "" && x_col %in% colnames(df))
    df[[x_col]] else seq_len(nrow(df))

  group_vals <- if (group_col != "" && group_col %in% colnames(df))
    as.character(df[[group_col]]) else NULL

  if (length(y_cols) > 1) {
    rep_mat <- sapply(y_cols, function(c) suppressWarnings(as.numeric(df[[c]])))
    y_mean  <- rowMeans(rep_mat, na.rm=TRUE)
    y_err   <- apply(rep_mat, 1, sd, na.rm=TRUE)   # sd() uses ddof=1 (N-1), matching Python
    y_reps  <- lapply(seq_len(ncol(rep_mat)), function(j) rep_mat[, j])
  } else {
    y_col  <- if (length(y_cols) == 1) y_cols[1] else NULL
    y_mean <- if (!is.null(y_col) && y_col %in% colnames(df))
      suppressWarnings(as.numeric(df[[y_col]])) else rep(NA_real_, nrow(df))
    y_reps <- NULL
    y_err  <- if (err_col != "" && err_col %in% colnames(df))
      suppressWarnings(as.numeric(df[[err_col]])) else rep(NA_real_, length(y_mean))
  }

  list(x_vals=x_vals, y_mean=y_mean, y_err=y_err, y_reps=y_reps, group_vals=group_vals)
}

# ── Shared helpers ─────────────────────────────────────────────────────────────

has_errors <- function(err_vec) any(!is.na(err_vec))

save_plot_file <- function(p, out_dir, plot_id, width, height) {
  out <- file.path(out_dir, paste0(plot_id, ".pdf"))
  ggsave(out, plot=p, device="pdf", width=width, height=height, units="in")
  message("    Saved -> ", basename(out))
}

cfg_labs <- function(p, cfg) {
  x_lbl <- if (cfg$x_label != "") cfg$x_label else NULL
  y_lbl <- if (cfg$y_label != "") cfg$y_label else NULL
  ttl   <- if (cfg$title   != "") cfg$title   else NULL
  p + labs(x=x_lbl, y=y_lbl, title=ttl)
}

# ── Bar chart ──────────────────────────────────────────────────────────────────

plot_bar <- function(xlsx_path, cfg, out_dir) {
  df <- load_sheet(xlsx_path, cfg)
  s  <- extract_series(df, cfg)

  n_x   <- length(unique(s$x_vals))
  n_g   <- if (!is.null(s$group_vals)) length(unique(s$group_vals)) else 1L
  fig_w <- max(3.5, n_x * 0.5 + n_g * 0.4 + 1.0)

  p <- if (is.null(s$group_vals)) simple_bar(s, cfg) else grouped_bar(s, cfg)
  save_plot_file(p, out_dir, cfg$plot_id, fig_w, 4)
}

simple_bar <- function(s, cfg) {
  n      <- length(s$y_mean)
  x_pos  <- seq_len(n)
  color  <- PALETTE[1]
  width  <- 0.6

  df_bar <- data.frame(
    x     = x_pos,
    y     = s$y_mean,
    y_err = s$y_err,
    label = as.character(s$x_vals),
    stringsAsFactors = FALSE
  )

  # Y-axis layout: respect user-specified limits; otherwise auto from data + errors
  err_safe  <- ifelse(is.na(df_bar$y_err), 0, df_bar$y_err)
  user_lo   <- parse_ylim(cfg$y_min)
  user_hi   <- parse_ylim(cfg$y_max)
  y_lo_data <- user_lo %||% 0
  y_hi_data <- user_hi %||% max(df_bar$y + err_safe, na.rm=TRUE)
  layout    <- make_axis_layout(y_lo_data, y_hi_data, user_lo=user_lo, user_hi=user_hi)

  p <- ggplot(df_bar, aes(x=x, y=y)) +
    geom_col(width=width, fill=color, color=NA, alpha=0.85) +
    scale_y_continuous(limits=layout$lims, expand=expansion(0), breaks=layout$breaks)

  if (has_errors(df_bar$y_err)) {
    p <- p + geom_errorbar(
      aes(ymin=y - y_err, ymax=y + y_err),
      width=0.15, linewidth=0.6, color="black"
    )
  }

  if (!is.null(s$y_reps)) {
    dot_offset <- width * 0.25   # slightly right of error bar center
    reps_rows  <- lapply(s$y_reps, function(vals) {
      data.frame(x=x_pos + dot_offset, value=vals, stringsAsFactors=FALSE)
    })
    reps_long <- do.call(rbind, reps_rows)
    reps_long <- reps_long[!is.na(reps_long$value), ]
    if (nrow(reps_long) > 0) {
      p <- p + geom_point(
        data=reps_long, aes(x=x, y=value),
        shape=21, fill="white", color="black",
        size=1.5, stroke=0.5, inherit.aes=FALSE
      )
    }
  }

  p <- p +
    scale_x_continuous(breaks=x_pos, labels=df_bar$label) +
    coord_cartesian(clip="off") +
    theme_pub()

  cfg_labs(p, cfg)
}

grouped_bar <- function(s, cfg) {
  unique_x  <- unique(s$x_vals)
  unique_g  <- unique(s$group_vals)
  n_g       <- length(unique_g)
  bar_width <- 0.7 / n_g
  x_int_map <- setNames(seq_along(unique_x), as.character(unique_x))

  bars_list <- list()
  dots_list <- list()

  for (gi in seq_along(unique_g)) {
    grp    <- unique_g[gi]
    mask   <- s$group_vals == grp
    color  <- PALETTE[((gi - 1L) %% length(PALETTE)) + 1L]
    x_int  <- x_int_map[as.character(s$x_vals[mask])]
    offset <- (gi - 1L - (n_g - 1L) / 2) * bar_width
    bar_x  <- x_int + offset

    bars_list[[gi]] <- data.frame(
      x     = bar_x,
      y     = s$y_mean[mask],
      y_err = s$y_err[mask],
      grp   = grp,
      color = color,
      stringsAsFactors = FALSE
    )

    if (!is.null(s$y_reps)) {
      dot_x <- bar_x
      for (rep_vals in s$y_reps) {
        dots_list[[length(dots_list) + 1L]] <- data.frame(
          x=dot_x, value=rep_vals[mask], color=color, stringsAsFactors=FALSE
        )
      }
    }
  }

  bars_df      <- do.call(rbind, bars_list)
  bars_df$grp  <- factor(bars_df$grp, levels=unique_g)
  fill_vals    <- setNames(PALETTE[seq_along(unique_g)], unique_g)

  # Y-axis layout
  err_safe  <- ifelse(is.na(bars_df$y_err), 0, bars_df$y_err)
  user_lo   <- parse_ylim(cfg$y_min)
  user_hi   <- parse_ylim(cfg$y_max)
  y_lo_data <- user_lo %||% 0
  y_hi_data <- user_hi %||% max(bars_df$y + err_safe, na.rm=TRUE)
  layout    <- make_axis_layout(y_lo_data, y_hi_data, user_lo=user_lo, user_hi=user_hi)

  p <- ggplot() +
    geom_col(
      data=bars_df, aes(x=x, y=y, fill=grp),
      width=bar_width, color=NA, alpha=0.85
    ) +
    scale_y_continuous(limits=layout$lims, expand=expansion(0), breaks=layout$breaks)

  if (has_errors(bars_df$y_err)) {
    p <- p + geom_errorbar(
      data=bars_df, aes(x=x, ymin=y - y_err, ymax=y + y_err),
      width=0.15, linewidth=0.6, color="black"
    )
  }

  if (length(dots_list) > 0) {
    dots_df <- do.call(rbind, dots_list)
    dots_df <- dots_df[!is.na(dots_df$value), ]
    if (nrow(dots_df) > 0) {
      # Shift dots to the right of error bar center
      dots_df$x <- dots_df$x + bar_width * 0.25
      p <- p + geom_point(
        data=dots_df, aes(x=x, y=value),
        shape=21, fill="white", color="black",
        size=1.5, stroke=0.5, inherit.aes=FALSE
      )
    }
  }

  p <- p +
    scale_x_continuous(breaks=seq_along(unique_x), labels=as.character(unique_x)) +
    scale_fill_manual(values=fill_vals) +
    coord_cartesian(clip="off") +
    theme_pub() +
    theme(legend.position="right")

  cfg_labs(p, cfg)
}

# ── Line chart ─────────────────────────────────────────────────────────────────

plot_line <- function(xlsx_path, cfg, out_dir) {
  df <- load_sheet(xlsx_path, cfg)
  plot_line_from_df(df, cfg, out_dir)
}

plot_line_from_df <- function(df, cfg, out_dir) {
  s     <- extract_series(df, cfg)
  x_num <- suppressWarnings(as.numeric(s$x_vals))

  x_range <- diff(range(x_num, na.rm=TRUE))
  eb_w    <- if (is.finite(x_range) && x_range > 0) x_range * 0.02 else 0.1

  # Y-axis layout from data range
  err_safe  <- ifelse(is.na(s$y_err), 0, s$y_err)
  y_vals    <- c(s$y_mean - err_safe, s$y_mean + err_safe)
  y_vals    <- y_vals[!is.na(y_vals) & is.finite(y_vals)]
  user_lo   <- parse_ylim(cfg$y_min)
  user_hi   <- parse_ylim(cfg$y_max)
  y_lo_data <- user_lo %||% (if (length(y_vals) > 0L) min(y_vals) else 0)
  y_hi_data <- user_hi %||% (if (length(y_vals) > 0L) max(y_vals) else 1)
  layout    <- make_axis_layout(y_lo_data, y_hi_data, user_lo=user_lo, user_hi=user_hi)

  if (!is.null(s$group_vals)) {
    unique_g  <- unique(s$group_vals)
    color_map <- setNames(PALETTE[seq_along(unique_g)], unique_g)

    df_base        <- data.frame(x=x_num, y=s$y_mean, y_err=s$y_err,
                                 group=s$group_vals, stringsAsFactors=FALSE)

    p <- ggplot(df_base, aes(x=x, y=y, color=group, group=group)) +
      geom_line(linewidth=1.0) +
      geom_point(size=2, position=position_identity())

    if (has_errors(s$y_err)) {
      p <- p + geom_errorbar(aes(ymin=y - y_err, ymax=y + y_err),
                             width=eb_w * 0.75, linewidth=0.6, color="black")
    }

    p <- p +
      scale_y_continuous(limits=layout$lims, expand=expansion(0), breaks=layout$breaks) +
      scale_color_manual(values=color_map) +
      coord_cartesian(clip="off") +
      theme_pub() +
      theme(legend.position="right")

  } else {
    df_base <- data.frame(x=x_num, y=s$y_mean, y_err=s$y_err)

    p <- ggplot(df_base, aes(x=x, y=y)) +
      geom_line(color=PALETTE[1], linewidth=1.0) +
      geom_point(color=PALETTE[1], size=2, position=position_identity())

    if (has_errors(s$y_err)) {
      p <- p + geom_errorbar(aes(ymin=y - y_err, ymax=y + y_err),
                             width=eb_w * 0.75, linewidth=0.6, color="black")
    }

    p <- p +
      scale_y_continuous(limits=layout$lims, expand=expansion(0), breaks=layout$breaks) +
      coord_cartesian(clip="off") + theme_pub()
  }

  if (cfg$x_scale == "log") p <- p + scale_x_log10()

  p <- cfg_labs(p, cfg)
  save_plot_file(p, out_dir, cfg$plot_id, 5, 4)
}

# ── Scatter + Linear regression ────────────────────────────────────────────────

plot_scatter <- function(xlsx_path, cfg, out_dir) {
  df <- load_sheet(xlsx_path, cfg)

  x_col  <- cfg$x_col
  y_cols <- trimws(strsplit(cfg$y_cols, ",")[[1]])
  y_col  <- if (length(y_cols) >= 1L && y_cols[1] != "") y_cols[1] else NULL

  x_raw <- if (x_col != "" && x_col %in% colnames(df))
    suppressWarnings(as.numeric(df[[x_col]])) else rep(NA_real_, nrow(df))
  y_raw <- if (!is.null(y_col) && y_col %in% colnames(df))
    suppressWarnings(as.numeric(df[[y_col]])) else rep(NA_real_, nrow(df))

  valid  <- !is.na(x_raw) & !is.na(y_raw)
  x_data <- x_raw[valid]
  y_data <- y_raw[valid]

  df_plot <- data.frame(x=x_data, y=y_data)

  # Y-axis layout
  user_lo <- parse_ylim(cfg$y_min)
  user_hi <- parse_ylim(cfg$y_max)
  y_lo_data <- user_lo %||% (if (length(y_data) > 0L) min(y_data) else 0)
  y_hi_data <- user_hi %||% (if (length(y_data) > 0L) max(y_data) else 1)
  layout  <- make_axis_layout(y_lo_data, y_hi_data, user_lo=user_lo, user_hi=user_hi)

  p <- ggplot(df_plot, aes(x=x, y=y)) +
    geom_point(color=PALETTE[1], size=2.5, alpha=0.9, shape=21,
               fill=PALETTE[1], stroke=0.5) +
    scale_y_continuous(limits=layout$lims, expand=expansion(0), breaks=layout$breaks) +
    coord_cartesian(clip="off") +
    theme_pub()

  if (length(x_data) >= 2L) {
    fit       <- lm(y ~ x, data=df_plot)
    slope     <- coef(fit)[2L]
    intercept <- coef(fit)[1L]
    r2        <- summary(fit)$r.squared

    x_seq   <- seq(min(x_data), max(x_data), length.out=200L)
    df_line <- data.frame(x=x_seq, y=slope * x_seq + intercept)

    p <- p + geom_line(data=df_line, aes(x=x, y=y),
                       color=PALETTE[2], linewidth=1.5, alpha=0.9,
                       inherit.aes=FALSE)

    sign_ch <- if (intercept >= 0) "+" else "-"
    # Format slope and intercept with 4 significant digits
    eq_str  <- sprintf("y = %s\u00B7x %s %s\nR\u00B2 = %.4f",
                       formatC(slope,          digits=4L, format="g"),
                       sign_ch,
                       formatC(abs(intercept), digits=4L, format="g"),
                       r2)
    p <- p + annotate("text",
                      x=-Inf, y=Inf,
                      hjust=-0.1, vjust=1.4,
                      label=eq_str,
                      size=FS$annot / PT,
                      family=FONT)
  }

  p <- cfg_labs(p, cfg)
  save_plot_file(p, out_dir, cfg$plot_id, 5, 4)
}

# ── Dose-response (log-scale line) ────────────────────────────────────────────

plot_dose_response <- function(xlsx_path, cfg, out_dir) {
  cfg        <- modifyList(cfg, list(x_scale="log"))
  df         <- load_sheet(xlsx_path, cfg)
  x_col      <- cfg$x_col
  if (x_col != "" && x_col %in% colnames(df)) {
    df <- df[suppressWarnings(as.numeric(df[[x_col]])) > 0, , drop=FALSE]
  }
  plot_line_from_df(df, cfg, out_dir)
}

# ── Bar-Line 복합 플롯 ──────────────────────────────────────────────────────────

parse_bar_line_series <- function(cfg) {
  split_pipe <- function(s) trimws(strsplit(s, "\\|")[[1L]])
  pad        <- function(lst, n) c(lst, rep("", max(0L, n - length(lst))))

  types      <- split_pipe(cfg$series_types)
  y_cols_all <- split_pipe(cfg$series_y_cols)
  err_cols   <- split_pipe(cfg$series_err_cols)
  y_labels   <- split_pipe(cfg$series_y_labels)
  names_s    <- split_pipe(cfg$series_names)
  n          <- length(types)

  y_cols_all <- pad(y_cols_all, n)
  err_cols   <- pad(err_cols,   n)
  y_labels   <- pad(y_labels,   n)
  names_s    <- pad(names_s,    n)

  lapply(seq_len(n), function(i) {
    yc <- trimws(strsplit(y_cols_all[i], ",")[[1L]])
    yc <- yc[yc != ""]
    list(type=tolower(types[i]), y_cols=yc, err_col=err_cols[i],
         y_label=y_labels[i], name=names_s[i])
  })
}

# Extract per-series data from df
extract_series_for_bar_line <- function(df, s) {
  y_cols  <- s$y_cols
  err_col <- trimws(s$err_col %||% "")

  if (length(y_cols) > 1L) {
    rep_mat <- sapply(y_cols, function(c) suppressWarnings(as.numeric(df[[c]])))
    list(
      y_mean = rowMeans(rep_mat, na.rm=TRUE),
      y_err  = apply(rep_mat, 1L, sd, na.rm=TRUE),
      y_reps = lapply(seq_len(ncol(rep_mat)), function(j) rep_mat[, j])
    )
  } else if (length(y_cols) == 1L) {
    list(
      y_mean = suppressWarnings(as.numeric(df[[y_cols[1L]]])),
      y_err  = if (err_col != "" && err_col %in% colnames(df))
        suppressWarnings(as.numeric(df[[err_col]])) else rep(NA_real_, nrow(df)),
      y_reps = NULL
    )
  } else NULL
}

plot_bar_line <- function(xlsx_path, cfg, out_dir) {
  df    <- load_sheet(xlsx_path, cfg)
  x_col <- cfg$x_col
  x_vals <- if (x_col != "" && x_col %in% colnames(df))
    df[[x_col]] else seq_len(nrow(df))
  x_int    <- seq_along(x_vals)
  x_labels <- as.character(x_vals)

  series_list <- parse_bar_line_series(cfg)
  n_series    <- length(series_list)
  n_right     <- max(0L, n_series - 1L)
  fig_w       <- max(5.0, length(x_vals) * 0.6 + 1.5 + n_right * 0.9)

  series_data <- lapply(series_list, extract_series_for_bar_line, df=df)

  # Compute data range [lo, hi] for each series (no panel padding — make_axis_layout handles that).
  # Includes replicate values so individual dots are never clipped.
  compute_data_range <- function(d, is_bar) {
    err_safe <- ifelse(is.na(d$y_err), 0, d$y_err)
    vals     <- c(d$y_mean - err_safe, d$y_mean + err_safe)
    if (!is.null(d$y_reps)) vals <- c(vals, unlist(d$y_reps))
    vals <- vals[!is.na(vals) & is.finite(vals)]
    if (length(vals) == 0L) return(c(0, 1))
    y_lo <- if (is_bar) 0 else min(vals)
    c(y_lo, max(vals))
  }

  # Parse per-series user-specified data limits (tick bounds) from Excel.
  parse_series_ylims <- function(col_name) {
    raw <- cfg[[col_name]]
    if (is.null(raw) || trimws(raw) == "") return(rep(list(NULL), n_series))
    parts <- trimws(strsplit(raw, "\\|")[[1L]])
    parts <- c(parts, rep("", max(0L, n_series - length(parts))))
    lapply(parts[seq_len(n_series)], parse_ylim)
  }
  user_ymins <- parse_series_ylims("series_y_mins")
  user_ymaxs <- parse_series_ylims("series_y_maxs")

  # data_range_list: [data_lo, data_hi] per series (user overrides applied)
  # layout_list: make_axis_layout result — $breaks and $lims (panel limits with top gap)
  data_range_list <- lapply(seq_len(n_series), function(i) {
    d    <- series_data[[i]]
    auto <- if (is.null(d)) c(0, 1) else
            compute_data_range(d, series_list[[i]]$type == "bar")
    lo   <- user_ymins[[i]] %||% auto[1L]
    hi   <- user_ymaxs[[i]] %||% auto[2L]
    c(lo, hi)
  })
  layout_list <- lapply(seq_len(n_series), function(i) {
    r <- data_range_list[[i]]
    make_axis_layout(r[1L], r[2L],
                     user_lo = user_ymins[[i]],
                     user_hi = user_ymaxs[[i]])
  })

  # Shared X coordinate system for all series panels — guarantees line dots land
  # exactly on tick marks when secondary panels are overlaid via gtable.
  n_x    <- length(x_int)
  x_lims <- c(0.5, n_x + 0.5)

  color_for <- function(i) PALETTE[((i - 1L) %% length(PALETTE)) + 1L]

  # ── Build data layer for each series ─────────────────────────────────────
  # Each plot shares the same x positions; Y-axis visibility controlled per role.
  make_data_plot <- function(i, show_left_yaxis, show_x_axis) {
    s      <- series_list[[i]]
    d      <- series_data[[i]]
    if (is.null(d)) return(NULL)
    color  <- color_for(i)
    layout <- layout_list[[i]]

    if (s$type == "bar") {
      bar_w <- 0.4
      df_s  <- data.frame(x=x_int, y_mean=d$y_mean, y_err=d$y_err)
      p <- ggplot(df_s, aes(x=x, y=y_mean)) +
        geom_col(width=bar_w, fill=color, color=NA, alpha=0.85) +
        scale_y_continuous(limits=layout$lims, expand=expansion(0), breaks=layout$breaks,
                           name=if (show_left_yaxis) s$y_label else NULL,
                           position="left")
      if (has_errors(d$y_err))
        p <- p + geom_errorbar(
          aes(ymin=pmax(layout$lims[1L], y_mean - y_err), ymax=y_mean + y_err),
          width=0.15, linewidth=0.6, color="black"
        )
      if (!is.null(d$y_reps)) {
        reps_long  <- do.call(rbind, lapply(d$y_reps, function(v)
          data.frame(x=x_int + bar_w * 0.25, value=v)))
        reps_long  <- reps_long[!is.na(reps_long$value), ]
        if (nrow(reps_long) > 0)
          p <- p + geom_point(data=reps_long, aes(x=x, y=value),
                              shape=21, fill="white", color="black",
                              size=1.5, stroke=0.5, inherit.aes=FALSE)
      }
    } else {
      df_s <- data.frame(x=x_int, y_mean=d$y_mean, y_err=d$y_err)
      p <- ggplot(df_s, aes(x=x, y=y_mean)) +
        geom_line(aes(group=1L), color=color, linewidth=1.0) +
        geom_point(color=color, size=2, position=position_identity()) +
        scale_y_continuous(limits=layout$lims, expand=expansion(0), breaks=layout$breaks,
                           name=if (show_left_yaxis) s$y_label else NULL,
                           position="left")
      if (has_errors(d$y_err))
        p <- p + geom_errorbar(aes(ymin=y_mean-y_err, ymax=y_mean+y_err),
                               width=0.15, linewidth=0.6, color="black")
    }

    x_lbl <- if (show_x_axis && cfg$x_label != "") cfg$x_label else NULL
    ttl   <- if (show_x_axis && cfg$title   != "") cfg$title   else NULL
    p <- p +
      scale_x_continuous(limits=x_lims, breaks=x_int, labels=x_labels,
                         expand=expansion(0), name=x_lbl) +
      labs(title=ttl) +
      coord_cartesian(clip="off") +
      theme_pub() +
      theme(legend.position="none")

    if (!show_x_axis)
      p <- p + theme(axis.text.x=element_blank(), axis.ticks.x=element_blank(),
                     axis.title.x=element_blank(), axis.line.x=element_blank())
    if (!show_left_yaxis)
      p <- p + theme(axis.text.y=element_blank(), axis.ticks.y=element_blank(),
                     axis.title.y=element_blank(), axis.line.y=element_blank(),
                     panel.background=element_rect(fill=NA, color=NA),
                     plot.background =element_rect(fill=NA, color=NA))
    p
  }

  # ── Build axis-only plot for each secondary series ────────────────────────
  # No data geom — just a right-side Y-axis with the correct scale/limits.
  make_axis_plot <- function(i) {
    s      <- series_list[[i]]
    layout <- layout_list[[i]]
    df_dummy <- data.frame(x=range(x_int), y=layout$lims)
    ggplot(df_dummy, aes(x=x, y=y)) +
      geom_blank() +
      scale_y_continuous(limits=layout$lims, expand=expansion(0),
                         breaks=layout$breaks, name=s$y_label, position="right") +
      scale_x_continuous() +
      coord_cartesian(clip="off") +
      theme_pub() +
      theme(
        panel.background    = element_rect(fill=NA, color=NA),
        plot.background     = element_rect(fill=NA, color=NA),
        plot.margin         = margin(0, 0, 0, 0),
        axis.line.x         = element_blank(), axis.ticks.x      = element_blank(),
        axis.text.x         = element_blank(), axis.title.x      = element_blank(),
        axis.line.y.left    = element_blank(), axis.ticks.y.left = element_blank(),
        axis.text.y.left    = element_blank(), axis.title.y.left = element_blank(),
        axis.line.y.right   = element_line(),
        axis.title.y.right  = element_text(size=FS$label, family=FONT,
                                           angle=90, margin=margin(l=8)),
        legend.position     = "none"
      )
  }

  # ── Assemble via gtable ───────────────────────────────────────────────────
  # Primary plot: full axes + data
  p1 <- make_data_plot(1L, show_left_yaxis=TRUE, show_x_axis=TRUE)

  # Embed legend into primary ggplot using alpha=0 ghost points so the legend
  # is positioned by ggplot2's own layout engine (avoids manual coordinate guessing).
  series_names <- sapply(series_list, function(s) s$name)
  if (any(series_names != "")) {
    ldf_leg <- data.frame(
      x    = rep(x_int[1L], n_series),
      y    = rep(0, n_series),
      name = factor(series_names, levels=series_names),
      stringsAsFactors = FALSE
    )
    # Bar type → filled square (shape=22), line type → filled circle (shape=21) with line
    leg_shapes    <- sapply(series_list, function(s) if (s$type == "bar") 22L else 21L)
    leg_linetypes <- sapply(series_list, function(s) if (s$type == "bar") 0L  else 1L)
    leg_sizes     <- sapply(series_list, function(s) if (s$type == "bar") 4.5 else 3.0)
    p1 <- p1 +
      geom_point(data=ldf_leg, aes(x=x, y=y, color=name),
                 size=3, alpha=0, inherit.aes=FALSE, show.legend=TRUE) +
      scale_color_manual(
        values = setNames(PALETTE[seq_len(n_series)], series_names),
        name   = NULL,
        guide  = guide_legend(override.aes=list(
          alpha    = 1,
          shape    = leg_shapes,
          linetype = leg_linetypes,
          size     = leg_sizes
        ))
      ) +
      theme(legend.position      = c(0.02, 0.98),
            legend.justification = c("left", "top"),
            legend.background    = element_rect(fill="white", linewidth=0.3,
                                                color="grey80"),
            legend.key           = element_blank())
  }

  g_main    <- ggplot_gtable(ggplot_build(p1))
  panel_pos <- g_main$layout[g_main$layout$name == "panel", ]

  # Insert new right-axis columns immediately after the panel (panel_pos$r),
  # so they are adjacent to the plot area — not separated by right-margin columns.
  insert_pos <- panel_pos$r

  for (i in seq(2L, n_series)) {
    if (is.null(series_data[[i]])) next

    # Overlay data panel (transparent background, no axes)
    p_data  <- make_data_plot(i, show_left_yaxis=FALSE, show_x_axis=FALSE)
    g_data  <- ggplot_gtable(ggplot_build(p_data))
    g_main  <- gtable::gtable_add_grob(
      g_main,
      g_data$grobs[[which(g_data$layout$name == "panel")]],
      t=panel_pos$t, l=panel_pos$l, b=panel_pos$b, r=panel_pos$r,
      clip="off", z=Inf, name=paste0("panel-sec-", i)
    )

    # Add right Y-axis columns right after panel (no gap from right margins)
    g_axis        <- ggplot_gtable(ggplot_build(make_axis_plot(i)))
    axis_r_layout <- g_axis$layout[g_axis$layout$name == "axis-r", ]
    ylab_r_layout <- g_axis$layout[g_axis$layout$name == "ylab-r", ]

    if (nrow(axis_r_layout) > 0L) {
      g_main <- gtable::gtable_add_cols(g_main, g_axis$widths[axis_r_layout$l], insert_pos)
      g_main <- gtable::gtable_add_grob(
        g_main,
        g_axis$grobs[[which(g_axis$layout$name == "axis-r")]],
        t=panel_pos$t, l=insert_pos+1L, b=panel_pos$b, r=insert_pos+1L,
        clip="off", name=paste0("axis-r-", i)
      )
      insert_pos <- insert_pos + 1L
    }
    if (nrow(ylab_r_layout) > 0L) {
      g_main <- gtable::gtable_add_cols(g_main, g_axis$widths[ylab_r_layout$l], insert_pos)
      g_main <- gtable::gtable_add_grob(
        g_main,
        g_axis$grobs[[which(g_axis$layout$name == "ylab-r")]],
        t=panel_pos$t, l=insert_pos+1L, b=panel_pos$b, r=insert_pos+1L,
        clip="off", name=paste0("ylab-r-", i)
      )
      insert_pos <- insert_pos + 1L
    }
  }

  # Legend is embedded in g_main. Render directly via grid (no cowplot wrapper)
  # to avoid extra viewport clipping that cowplot::ggdraw() adds.
  out <- file.path(out_dir, paste0(cfg$plot_id, ".pdf"))
  grDevices::pdf(out, width=fig_w, height=4)
  grid::grid.newpage()
  grid::grid.draw(g_main)
  grDevices::dev.off()
  message("    Saved -> ", basename(out))
}

# ── Dispatch ───────────────────────────────────────────────────────────────────

DISPATCH <- list(
  bar           = plot_bar,
  line          = plot_line,
  scatter       = plot_scatter,
  dose_response = plot_dose_response,
  bar_line      = plot_bar_line
)

# ── Entry point ────────────────────────────────────────────────────────────────

main <- function(xlsx_path_str) {
  # 내부 연산(gtable, cowplot 등)이 기본 디바이스(Rplots.pdf)를 만들지 않도록
  # options(device=)를 null 디바이스로 교체한다.
  old_device <- options(device=function(...) grDevices::pdf(nullfile()))
  on.exit(options(old_device), add=TRUE)

  xlsx_path <- normalizePath(xlsx_path_str, mustWork=TRUE)
  out_dir   <- dirname(xlsx_path)

  message("\nReading: ", basename(xlsx_path))

  configs <- load_configs(xlsx_path)
  message("Found ", length(configs), " plot(s) in '__plots__' sheet.\n")

  ok <- 0L; err <- 0L
  for (cfg in configs) {
    plot_id   <- cfg$plot_id
    plot_type <- tolower(cfg$plot_type)
    message("  [", plot_type, "] ", plot_id)

    fn <- DISPATCH[[plot_type]]
    if (is.null(fn)) {
      message("    Warning: unknown plot_type '", plot_type, "', skipping.")
      next
    }

    tryCatch({
      fn(xlsx_path, cfg, out_dir)
      ok <- ok + 1L
    }, error=function(e) {
      message("    Error: ", conditionMessage(e))
      err <<- err + 1L
    })
  }

  message("\nDone - ", ok, " succeeded, ", err, " failed.")
}

args <- commandArgs(trailingOnly=TRUE)
if (length(args) >= 1L) {
  main(args[1L])
}
