# Plot Automation Pipeline — bar_line 다축 스타일 통일

## Context

`bar_line` 다축 그래프에서 두 가지 추가 수정:
1. **spine 색상 제거**: 우측 Y축 선(spine)과 tick mark가 series 색으로 표시됨 → 전부 검정
2. **tick 위치 정렬**: 여러 우측 축이 각자 독립적으로 tick을 생성하여 수평 위치가 어긋남 → 모든 축이 동일한 비율 위치에 tick을 갖도록 통일

수정 대상: `/Users/hojoon/workspace/Plot_claude_automation/plot.py` — `plot_bar_line` 함수

---

## 1. spine 색상 제거 (현재 line 527–530)

현재:
```python
if i == 0:
    ax.spines["left"].set_color(color)
else:
    ax.spines["right"].set_color(color)
```

→ 이 블록 전체 삭제. spine 기본색은 검정이므로 별도 설정 불필요.

## 2. Y축 tick 정렬 (series 루프 종료 후 삽입)

**원리**: `twinx()`로 생성된 모든 축은 동일한 figure 공간을 공유함.
각 축에서 `np.linspace(ymin, ymax, n)` 으로 tick을 설정하면 → 모든 축에서 비율 위치 0%, 25%, 50%, 75%, 100%에 tick이 놓임 → 픽셀 위치가 일치.

series 루프 종료 직후, `# 공통 X축 스타일` 앞에 삽입:

```python
# 모든 Y축 tick 위치 정렬: linspace로 동일 비율 위치 → 시각적 정렬 보장
n_ticks = 5
for ax in axes:
    ymin, ymax = ax.get_ylim()
    ax.set_yticks(np.linspace(ymin, ymax, n_ticks))
    for lbl in ax.get_yticklabels():
        lbl.set_fontfamily(FONT)
```

기존 루프 내부의 `for lbl in ax.get_yticklabels(): lbl.set_fontfamily(FONT)` 는 삭제
(tick을 재설정하면 label 객체가 교체되므로, 위 블록에서 일괄 처리).

---

## 최종 결과

- 모든 Y축 선·tick mark·숫자: 검정
- 우측 축이 2개든 3개든 tick 수평 위치 완전 일치
- 숫자값은 각 축의 실제 스케일에 맞춰 다름 (당연)

## 검증

```bash
.venv/bin/python3 plot.py example_template.xlsx
# 8 succeeded, 0 failed
```
