from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

OUT = Path(__file__).with_name("KNN_실습_보고서.docx")
SC = Path("/private/tmp/claude-501/-Users-seohyeokin-workspace-skala-MachineLearning/6cb4b9c6-686b-44fa-8232-57d9fb82d3e1/scratchpad")
FONT = "Apple SD Gothic Neo"


def set_font(run, size=9, bold=False, color=None, italic=False):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)


def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_table_borders(table, color="D9D9D9"):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:color"), color)
        borders.append(el)
    tbl_pr.append(borders)


def para(doc, text, size=9, bold=False, after=3, before=0, align=None, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.line_spacing = 1.08
    if align:
        p.alignment = align
    set_font(p.add_run(text), size=size, bold=bold, italic=italic)
    return p


def heading(doc, text, size=12):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(text), size=size, bold=True, color=(31, 78, 120))
    return p


def subheading(doc, text, size=9.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(text), size=size, bold=True)
    return p


def table(doc, headers, rows, widths, header_fill="1F4E78", fsize=8.3):
    t = doc.add_table(rows=1, cols=len(headers))
    t.autofit = False
    t.style = "Table Grid"
    set_table_borders(t)
    for i, text in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.width = Inches(widths[i])
        shade(cell, header_fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        set_font(p.add_run(text), size=fsize, bold=True, color=(255, 255, 255))
    for r_idx, values in enumerate(rows):
        cells = t.add_row().cells
        for i, text in enumerate(values):
            cells[i].width = Inches(widths[i])
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if r_idx % 2 == 1:
                shade(cells[i], "EEF4F8")
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            set_font(p.add_run(text), size=fsize)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


doc = Document()
section = doc.sections[0]
section.top_margin = section.bottom_margin = Inches(0.5)
section.left_margin = section.right_margin = Inches(0.6)

styles = doc.styles
styles["Normal"].font.name = FONT
styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
styles["Normal"].font.size = Pt(9)

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_after = Pt(2)
set_font(title.add_run("[머신러닝 및 딥러닝] 실습 #2 : KNN 결과 보고서"), size=15, bold=True)
para(doc, "데이터셋: Breast Cancer Wisconsin (Diagnostic) — 569 samples, 30 features, Train/Test = 80/20 (Validation 없이 단순화)",
     size=8.5, after=8, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)

# ===================== (1) 개별 실험 =====================
heading(doc, "(1) 개별 실험", size=11)
para(doc, "설정값 근거 — Train/Test=80/20·random_state=42: 과제 지침대로 Validation 없이 단순화, 재현성 확보를 위해 seed 고정. "
          "K 탐색범위 1~20(홀수 우선): 1은 최댓값 과적합 기준선, 20은 전체(569)의 약 3.5%로 충분히 큰 과소적합 기준선; 짝수는 "
          "이진분류 동점 방지를 위해 후순위로 배치. Scaling은 StandardScaler(평균0/분산1) 사용: KNN이 거리 기반이라 특성 간 "
          "스케일 차이의 영향을 가장 표준적으로 제거하는 방법이기 때문. metric은 minkowski 계열의 대표값인 euclidean(p=2)·"
          "manhattan(p=1) 2종, weights는 uniform·distance 2종만 비교: 과제에서 요구한 '필요 시 추가 실험' 범위에서 해석 가능한 "
          "최소 조합으로 한정.", size=7.8, after=5, italic=True)

subheading(doc, "① K / Scaling 비교 (Test Accuracy)", size=9)
table(
    doc,
    ["K", "1", "3", "5", "7", "9", "11"],
    [
        ["Scaling 전"] + ["0.8947", "0.9211", "0.9123", "0.9123", "0.9211", "0.9474"],
        ["Scaling 후"] + ["0.9298", "0.9386", "0.9561", "0.9561", "0.9474", "0.9474"],
        ["Train(후)"] + ["1.0000", "0.9846", "0.9780", "0.9714", "0.9736", "0.9692"],
    ],
    widths=[0.75] + [0.65] * 6,
    fsize=7.3,
)
doc.add_picture(str(SC / "chart_combined.png"), width=Inches(6.2))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

subheading(doc, "② metric × weights 추가 실험 (Test Accuracy)", size=9)
table(
    doc,
    ["K", "euclid/uniform", "euclid/distance", "manhattan/uniform", "manhattan/distance"],
    [
        ["3", "0.9386", "0.9386", "0.9649 ★", "0.9649"],
        ["5", "0.9561", "0.9561", "0.9561", "0.9561"],
        ["9", "0.9474", "0.9474", "0.9561", "0.9561"],
    ],
    widths=[0.4, 1.35, 1.35, 1.4, 1.4],
    fsize=7.3,
)

subheading(doc, "③ 해석 및 영향력 순위", size=9)
para(doc, "K=1은 과적합(Train 1.00/Test 0.93), K=20은 과소적합 경향이며 K=5~7에서 Test 최고(0.9561)를 기록해 전형적 U자형 곡선을 보임. "
          "Scaling은 전 구간 성능을 높였고(최대 +4.39%p, 특성 간 값 크기 차이로 거리 계산이 왜곡되기 때문), metric을 manhattan으로 "
          "바꾸면 K=3에서 0.9649로 K 단독 튜닝보다 더 좋아짐. Euclidean은 차이를 제곱(∑(x-y)²)해 큰 값을 가진 특성이 거리를 지배하는 "
          "반면, Manhattan은 절댓값 합(∑|x-y|)이라 30개 특성 중 튀는 값 1~2개에 결과가 휘둘리는 정도가 작음 — 이 효과는 여러 이웃을 "
          "평균 내는 K가 클수록(5, 9) 희석되고, 이웃 소수만 보는 K=3처럼 작을 때 가장 뚜렷하게 나타남. weights는 uniform/distance 간 "
          "차이가 거의 없었음.",
     after=3)
table(
    doc,
    ["순위", "요인", "조작", "변화폭"],
    [
        ["1", "Scaling", "미적용→적용", "+4.39%p"],
        ["2", "metric", "euclidean→manhattan", "+2.63%p"],
        ["3", "K", "1→5 / 5→20", "+2.63%p / -1.75%p"],
        ["4", "weights", "uniform↔distance", "0.00%p"],
    ],
    widths=[0.4, 0.8, 1.7, 1.4],
    fsize=7.3,
)
para(doc, "🏆 Best: StandardScaler + manhattan + uniform + K=3 (Test Acc 0.9649). 영향력: Scaling > metric ≈ K > weights.",
     bold=True, after=2)

# ===================== (2) 조별 토의 =====================
heading(doc, "(2) 조별 토의", size=11)
table(
    doc,
    ["항목", "내용"],
    [
        ["Best Model / 이유", "[조원 결과 취합 후 기입 — 본인: K=3·manhattan·uniform·Scaled, Acc 0.9649. "
                             "선정 기준: Test Acc + Train-Test 격차 + Recall 종합]"],
        ["Bias/Variance 비교", "K↓ = Low Bias/High Variance(과적합), K↑ = High Bias/Low Variance(과소적합) — "
                              "조원별 채택 K의 Train-Test 격차로 비교, 조원 결과 취합 후 기입"],
        ["주요 차이·원인", "[조원 간 최적 K·metric·scaling 차이와 원인(분할 seed, 탐색 범위 등) — 논의 후 기입]"],
    ],
    widths=[1.3, 5.9],
    fsize=7.8,
)

# ===================== (3) 개인 시사점 =====================
heading(doc, "(3) 개인 시사점", size=11)
para(doc, "KNN 성능은 K 자체보다 전처리·거리척도 선택이 더 크게 좌우했다(Scaling +4.4%p > metric +2.6%p > K +2.6%p > weights 0%p). "
          "'최적 K'는 고정값이 아니라 metric·scaling 조합에 따라 달라진다는 점(euclidean-K5 vs manhattan-K3)이 인상적이었고, "
          "Accuracy 단일 지표보다 Train-Test 격차·Recall을 함께 봐야 신뢰할 수 있는 모델을 고를 수 있음을 체감했다.")

doc.save(OUT)
print(f"Saved: {OUT}")
