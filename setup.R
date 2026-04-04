#!/usr/bin/env Rscript
# setup.R — plot.R 실행에 필요한 R 패키지를 설치한다.
#
# 사용법:
#   Rscript setup.R

required <- c("readxl", "ggplot2", "scales", "cowplot")
missing  <- required[!required %in% installed.packages()[, "Package"]]

if (length(missing) > 0) {
  message("Installing missing packages: ", paste(missing, collapse=", "))
  install.packages(missing, repos="https://cloud.r-project.org")
} else {
  message("All required packages are already installed.")
}

# 설치 확인
for (pkg in required) {
  if (requireNamespace(pkg, quietly=TRUE)) {
    message("  [OK] ", pkg)
  } else {
    message("  [FAIL] ", pkg, " — please install manually.")
  }
}

message("\nSetup complete. Run with: Rscript plot.R <experiment>.xlsx")
