# Plot Automation

Excel 파일 하나로 publication-quality PDF/SVG 그래프를 자동 생성하는 도구.

---

## 설치

### R 방식 (권장 — PDF, 고품질)

**요구사항**: R 4.0+ ([r-project.org](https://www.r-project.org/)에서 설치)

```bash
# 최초 1회만: R 패키지 설치
Rscript setup.R

# Excel 파일에서 PDF 그래프 생성
Rscript plot.R <experiment>.xlsx
```

### Python 방식 (레거시 — SVG/PDF)

**요구사항**: Python 3.9 이상 ([python.org](https://www.python.org/downloads/)에서 설치)

```bash
# 가상 환경 생성
python3 -m venv .venv

# 패키지 설치 (Mac/Linux)
.venv/bin/pip install -r requirements.txt

# 패키지 설치 (Windows)
.venv\Scripts\pip install -r requirements.txt
```

## 실행

### R 방식
```bash
Rscript plot.R <experiment>.xlsx
```

### Python 방식
```bash
# Mac/Linux
.venv/bin/python3 plot.py <experiment>.xlsx              # PDF 출력 (기본값)
.venv/bin/python3 plot.py <experiment>.xlsx --format svg # SVG 출력

# Windows
.venv\Scripts\python plot.py <experiment>.xlsx
.venv\Scripts\python plot.py <experiment>.xlsx --format svg
```

출력 파일은 Excel 파일과 같은 디렉터리에 `plot_id.pdf` 또는 `plot_id.svg` 이름으로 저장된다.

---

## Excel 파일 구성

Excel 파일은 반드시 `__plots__` 시트를 포함해야 한다. 각 행이 그래프 1개를 정의한다.
데이터는 별도 시트에 wide format으로 작성한다.

`example_template.xlsx`의 **"사용법" 시트**에 전체 컬럼 규약과 예시가 정리되어 있다.

### `__plots__` 시트 — 공통 컬럼

| 컬럼 | 필수 | 설명 |
|------|------|------|
| `plot_id` | 필수 | 출력 파일명 (확장자 제외) |
| `plot_type` | 필수 | `bar` / `line` / `scatter` / `dose_response` / `bar_line` |
| `data_sheet` | 필수 | 데이터 시트명 |
| `data_start_row` | 선택 | 헤더 행 번호(1-indexed). 생략 시 1행 |
| `x_col` | 필수 | X축 컬럼명 |
| `y_cols` | 필수 | Y값 컬럼명. 콤마 구분 시 triplicate → 자동 mean/std |
| `err_col` | 선택 | 사전 계산된 오차 컬럼 (`y_cols`가 단일 mean일 때) |
| `group_col` | 선택 | 그룹/색상 구분 컬럼 (grouped bar, multi-line) |
| `x_label` | 선택 | X축 레이블 |
| `y_label` | 선택 | Y축 레이블 |
| `title` | 선택 | 그래프 제목 |
| `x_scale` | 선택 | `log` 입력 시 로그 스케일 (기본: linear) |
| `y_min` | 선택 | Y축 눈금 하한 (자동 계산 시 생략) |
| `y_max` | 선택 | Y축 눈금 상한 (자동 계산 시 생략) |

### `bar_line` 전용 추가 컬럼

하나의 차트에 bar와 line이 공존하는 다축 복합 그래프. 각 series가 독립 Y축을 가진다.
Series 1 → 좌측 Y축, Series 2 → 첫 번째 우측 Y축, Series 3+ → 추가 우측 축 (65pt 간격).

| 컬럼 | 설명 | 예시 (3 series) |
|------|------|-----------------|
| `series_types` | 파이프(`\|`) 구분 series 타입 | `bar\|line\|line` |
| `series_y_cols` | 파이프 구분 Y 컬럼 (콤마로 replicate 지정) | `r1,r2,r3\|OD_mean\|pH` |
| `series_err_cols` | 파이프 구분 오차 컬럼 (replicate 사용 시 빈칸) | `\|OD_std\|` |
| `series_y_labels` | 파이프 구분 Y축 레이블 | `Indican (mM)\|OD600\|pH` |
| `series_names` | 파이프 구분 범례 이름 | `Indican\|OD600\|pH` |
| `series_y_mins` | 파이프 구분 각 series Y축 눈금 하한 | `0\|0\|6.5` |
| `series_y_maxs` | 파이프 구분 각 series Y축 눈금 상한 | `3\|1.5\|8` |

---

## 출력 스타일

**R 방식 (plot.R)**:
- 포맷: PDF, 투명 배경, publication-quality
- 폰트: Arial (또는 Helvetica/sans 폴백)
- 색상 팔레트: 8색 (muted natural 톤)
- Bar chart: 각 bar 옆에 개별 데이터 포인트(dot) 표시 (흰 채움, 검정 테두리)
- Y축 눈금: 자동 계산 또는 Excel로 직접 지정 가능 (`y_min`, `y_max`, `series_y_mins`, `series_y_maxs`)
- 모든 Y축: 상단 눈금이 panel 내 동일 높이에 정렬 (다축 차트에서 시각적 일관성)

**Python 방식 (plot.py)** (레거시):
- 포맷: SVG 또는 PDF 선택 가능, 투명 배경
- 폰트: Arial 고정
- 색상 팔레트: 8색
- Bar chart: 각 bar 위에 개별 데이터 포인트(dot) 오버레이
- 모든 텍스트(축 레이블, 숫자, 제목): 검정

**공통**:
- 9개 이상 조건이 필요한 경우 소스 코드 상단의 `PALETTE` 리스트에 색 추가

---

## 파일 구성

```
plot.R                   # R 메인 스크립트 (권장)
setup.R                  # R 패키지 설치 스크립트
plot.py                  # Python 메인 스크립트 (레거시)
requirements.txt         # Python 패키지 목록
example_template.xlsx    # 사용 예시 및 템플릿
  ├── 사용법             # 전체 사용 가이드
  ├── __plots__          # 예시 plot 설정
  └── data_*             # 예시 데이터 시트
CLAUDE.md                # Claude Code용 가이드
README.md                # 이 파일
```
