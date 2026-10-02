"""실습 #3 Clustering 보고서 생성. 실행: python3 build_clustering_report.py"""
import tempfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from sklearn.cluster import DBSCAN, AgglomerativeClustering, KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import calinski_harabasz_score, silhouette_score
from sklearn.preprocessing import StandardScaler
from scipy.cluster.hierarchy import dendrogram, linkage

OUT = Path(__file__).with_name("Clustering_실습_보고서.docx")
SC = Path(tempfile.mkdtemp())
FONT = "Apple SD Gothic Neo"
SPEND = ["Fresh", "Milk", "Grocery", "Frozen", "Detergents_Paper", "Delicassen"]

# ---------------------------------------------------------------- 실험 재현
from ucimlrepo import fetch_ucirepo

wc = fetch_ucirepo(id=292)
df = wc.data.features.join(wc.data.targets)
X = StandardScaler().fit_transform(np.log1p(df[SPEND].values.astype(float)))

ks = list(range(2, 9))
inertia = [KMeans(k, random_state=42, n_init=10).fit(X).inertia_ for k in ks]
sils = [silhouette_score(X, KMeans(k, random_state=42, n_init=10).fit_predict(X)) for k in ks]

km_lab = KMeans(4, random_state=42, n_init=10).fit_predict(X)
ward_lab = AgglomerativeClustering(n_clusters=4, linkage="ward").fit_predict(X)
db_lab = DBSCAN(eps=0.8, min_samples=6).fit_predict(X)

# k=3 대안과의 수치 비교용
km_lab3 = KMeans(3, random_state=42, n_init=10).fit_predict(X)
sil3, sil4 = silhouette_score(X, km_lab3), silhouette_score(X, km_lab)
ch3, ch4 = calinski_harabasz_score(X, km_lab3), calinski_harabasz_score(X, km_lab)
size3 = pd.Series(km_lab3).value_counts().sort_values(ascending=False)
size4 = pd.Series(km_lab).value_counts().sort_values(ascending=False)

Z = linkage(X, method="ward")
merge_h = Z[:, 2]  # ascending merge heights; last k-1'th value = height forming k clusters from k+1

fig, ax = plt.subplots(figsize=(9.2, 2.3))
dendrogram(Z, truncate_mode="lastp", p=24, leaf_rotation=90, leaf_font_size=6, ax=ax,
           color_threshold=merge_h[-3])
ax.axhline(merge_h[-3], ls="--", c="crimson", lw=1)
ax.set_title("Dendrogram (ward, log+scaled) — dashed line = k=4 cut", fontsize=9)
ax.set_ylabel("Distance", fontsize=8); ax.tick_params(labelsize=6)
plt.tight_layout(); plt.savefig(SC / "fig0.png", dpi=150); plt.close()

fig, ax = plt.subplots(1, 2, figsize=(9, 2.6))
ax[0].plot(ks, inertia, "bo-"); ax[0].set_title("Elbow (log+scaled)", fontsize=9)
ax[1].plot(ks, sils, "ro-"); ax[1].set_title("Silhouette", fontsize=9)
for a in ax:
    a.set_xlabel("k", fontsize=8); a.grid(alpha=.3); a.axvline(4, ls="--", c="gray")
    a.tick_params(labelsize=7)
plt.tight_layout(); plt.savefig(SC / "fig1.png", dpi=150); plt.close()

P = PCA(2).fit_transform(X)
fig, ax = plt.subplots(1, 3, figsize=(10, 2.5))
for a, (t, lab) in zip(ax, [("K-Means k=4", km_lab), ("Hierarchical-ward k=4", ward_lab),
                            ("DBSCAN eps=0.8 (noise=262)", db_lab)]):
    a.scatter(P[:, 0], P[:, 1], c=lab, cmap="viridis", s=10, alpha=.8)
    a.set_title(t, fontsize=8); a.grid(alpha=.3); a.tick_params(labelsize=6)
plt.tight_layout(); plt.savefig(SC / "fig2.png", dpi=150); plt.close()

df["cluster"] = km_lab
prof = df.groupby("cluster")[SPEND].mean().round(0).astype(int)
ratio = (df.groupby("cluster")[SPEND].mean() / df[SPEND].mean()).round(2)
size = df.groupby("cluster").size()
retail = (df.groupby("cluster").Channel.apply(lambda s: (s == 2).mean()) * 100).round(0).astype(int)

# ---------------------------------------------------------------- docx 헬퍼
def set_font(run, size=9, bold=False, color=None, italic=False):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size); run.bold = bold; run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)


def shade(cell, fill):
    shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def set_table_borders(table, color="D9D9D9"):
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "6"); el.set(qn("w:color"), color)
        borders.append(el)
    table._tbl.tblPr.append(borders)


def para(doc, text, size=9, bold=False, after=3, before=0, align=None, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.line_spacing = 1.05
    if align:
        p.alignment = align
    set_font(p.add_run(text), size=size, bold=bold, italic=italic)
    return p


def heading(doc, text, size=11):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(text), size=size, bold=True, color=(31, 78, 120))
    return p


def subheading(doc, text, size=9):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(text), size=size, bold=True)
    return p


def table(doc, headers, rows, widths, fsize=7.3):
    t = doc.add_table(rows=1, cols=len(headers))
    t.autofit = False; t.style = "Table Grid"
    set_table_borders(t)
    for i, text in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.width = Inches(widths[i]); shade(cell, "1F4E78")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        set_font(p.add_run(text), size=fsize, bold=True, color=(255, 255, 255))
    for r_idx, values in enumerate(rows):
        cells = t.add_row().cells
        for i, text in enumerate(values):
            cells[i].width = Inches(widths[i])
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if r_idx % 2 == 1:
                shade(cells[i], "EEF4F8")
            p = cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            set_font(p.add_run(text), size=fsize)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


def pic(doc, name, width):
    doc.add_picture(str(SC / name), width=Inches(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].paragraph_format.space_after = Pt(2)


# ---------------------------------------------------------------- 문서
doc = Document()
s = doc.sections[0]
s.top_margin = s.bottom_margin = Inches(0.45)
s.left_margin = s.right_margin = Inches(0.55)
doc.styles["Normal"].font.name = FONT
doc.styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
doc.styles["Normal"].font.size = Pt(9)

title = doc.add_paragraph(); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_after = Pt(2)
set_font(title.add_run("[머신러닝 및 딥러닝] 실습 #3 : Clustering 결과 보고서"), size=14.5, bold=True)
para(doc, "데이터셋: Wholesale Customers (UCI id=292) — 440 samples / 8 features · 도매업체 연간 품목별 구매액",
     size=8.3, after=6, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)

heading(doc, "(1) 개별 실험")
para(doc, "설정값 근거 — 클러스터링 변수는 6개 소비 금액 변수만 사용(Channel·Region은 이미 알려진 라벨이므로 검증·프로파일링용으로 분리). "
          "6개 변수의 왜도가 2.6~11.2로 극단적 우편향이라 log1p 후 StandardScaler를 적용(거리 기반 알고리즘에서 소수 대형 거래처가 "
          "거리를 지배하는 것을 막기 위함). k는 Elbow·Silhouette·Dendrogram 절단 거리를 함께 보고 2~8 범위 탐색, KMeans는 n_init=10·"
          "random_state=42로 초기값 의존성과 재현성을 통제. DBSCAN의 eps는 6-NN 거리 분포(k-distance)에서 후보를 잡아 0.8~1.8 grid 탐색, "
          "min_samples=6(변수 수+1).",
     size=7.6, after=4, italic=True)

subheading(doc, "① 전처리 비교 — 지표가 좋은 모델 ≠ 쓸모 있는 세분화 (KMeans k=4)")
table(doc, ["전처리", "Silhouette", "CH Index", "클러스터 크기", "최대 클러스터 비중", "판정"],
      [["원본(raw)", "0.390", "207.7", "277/95/58/10", "63.0%", "금액 큰 Fresh가 지배"],
       ["Scaling만", "0.348", "145.8", "315/110/10/5", "71.6%", "여전히 1개 군집에 쏠림"],
       ["log + Scaling", "0.188", "131.3", "132/129/118/61", "30.0%", "채택 — 4개 군집이 고르게 분산·해석 가능"]],
      widths=[0.95, 0.8, 0.75, 1.15, 1.1, 2.4])

subheading(doc, "② 클러스터 수(k) 탐색 — Dendrogram + Elbow + Silhouette")
pic(doc, "fig0.png", 6.3)
pic(doc, "fig1.png", 5.9)
para(doc, "Dendrogram(ward) 기준 병합 거리는 k=2→3에서 24.45, k=3→4에서 20.47, k=4→5에서 16.47로 k=3/k=4 절단 차이(3.98)가 "
          "k=4/k=5 차이(4.00)와 거의 같아 '3개에서 멈춰야 할' 뚜렷한 계단이 없다. k=3·k=4를 KMeans로 직접 비교한 수치는 아래와 같다.",
     size=7.5, after=2)
table(doc, ["k", "Silhouette", "CH Index", "클러스터 크기", "최대 비중", "판정"],
      [[f"k=3", f"{sil3:.3f}", f"{ch3:.1f}", "/".join(str(v) for v in size3), f"{size3.max()/size3.sum()*100:.1f}%",
        "지표 우세하나 대형/소형 거래처 차이가 뭉뚱그려짐"],
       [f"k=4", f"{sil4:.3f}", f"{ch4:.1f}", "/".join(str(v) for v in size4), f"{size4.max()/size4.sum()*100:.1f}%",
        "채택 — 지표 소폭 하락, 세분화 정보량 더 큼"]],
      widths=[0.6, 0.95, 0.85, 1.3, 0.95, 2.55])
para(doc, f"Silhouette·CH 모두 k=3({sil3:.3f}/{ch3:.1f})이 k=4({sil4:.3f}/{ch4:.1f})보다 높지만, 하락 폭이 완만하고(-0.071/-21.5) "
          "병합 거리상 k=3에서 멈춰야 할 근거도 없다. 이 정도 지표 손해로 ⑤처럼 같은 채널 안 대형/소형 거래처를 구분할 수 있다는 점을 "
          "우선해, 수치를 비교한 결과 k=4를 최종 채택했다.",
     size=7.5, after=2)

subheading(doc, "③ Clustering Method 비교 (log+scaled, k=4)")
table(doc, ["Method", "설정", "Silhouette", "CH", "클러스터 크기", "Noise", "해석"],
      [["K-Means", "k=4, n_init=10", "0.188", "131.3", "61/129/132/118", "0", "★ 최종 채택"],
       ["Hierarchical", "ward, k=4", "0.202", "108.6", "53/208/125/54", "0", "KMeans와 68.0% 일치(최적정렬)"],
       ["Hierarchical", "average, k=4", "0.498", "8.4", "437/1/1/1", "0", "singleton 분리(체이닝)"],
       ["Hierarchical", "complete, k=4", "0.260", "38.9", "391/33/11/5", "0", "불균형"],
       ["DBSCAN", "eps=0.8, min_s=6", "0.324", "141.3", "88/83/7", "262", "60%가 noise 처리"],
       ["DBSCAN", "eps=1.2~1.8", "-", "-", "360~418", "80→22", "전부 1개 군집으로 병합"]],
      widths=[0.85, 1.15, 0.75, 0.6, 1.0, 0.5, 2.3])
pic(doc, "fig2.png", 6.3)
para(doc, "합성 데이터 실험에서는 Moon 형태에 DBSCAN이 압도적이었으나, 실제 판매 데이터는 6차원에서 밀도가 연속적으로 이어져 "
          "eps를 조금만 키워도 전부 한 덩어리, 조금만 줄이면 60%가 noise가 되는 밀도 임계가 존재하지 않았다. 즉 데이터가 "
          "'명확히 분리된 덩어리'가 아니라 '연속 스펙트럼'일 때는 군집 수를 강제 지정하는 K-Means/Hierarchical이 적합하다. "
          "Hierarchical-average는 Silhouette 0.498로 1위지만 437/1/1/1로 쪼개진 결과이며, 이는 ①의 Scaling만 적용한 경우와 함께 "
          "'지표 최댓값이 곧 좋은 세분화는 아니다'를 다시 보여준다. KMeans와 Ward의 라벨 일치도(최적정렬 기준 68.0%)는 k=3일 때(74.8%)보다 "
          "낮아졌는데, 이는 k가 늘수록 두 알고리즘의 경계 판단(centroid 기반 vs 분산 최소화 병합)이 갈리는 지점이 늘어나기 때문 — "
          "세분화를 더 정교하게 할수록 method 선택의 영향력도 커진다는 것을 확인할 수 있다.", size=7.6, after=3)

subheading(doc, "④ 변수 구성 비교 (KMeans k=4)")
table(doc, ["변수 구성", "Silhouette", "CH", "크기", "Channel 일치도", "해석"],
      [["6개 소비 변수", "0.188", "131.3", "61/129/132/118", "88.0%", "구매 패턴 기반 세분화(채택)"],
       ["8개 전체(+Channel/Region)", "0.212", "124.6", "142/81/104/113", "97.3%", "Channel이 군집을 지배 → 신규 정보 없음"]],
      widths=[1.7, 0.8, 0.6, 1.0, 1.0, 2.1])

subheading(doc, "⑤ Customer Profiling — K-Means k=4 (원본 스케일 평균, 괄호는 전체평균 대비 배수)")
rows = []
names = {0: "그로서리·세제 중심 소형 소매형", 1: "신선·냉동 중심 요식업(Horeca)",
         2: "저지출 소형 요식업(Horeca)", 3: "대형 고액 종합구매 소매/유통형"}
for c in [3, 1, 0, 2]:
    rows.append([f"C{c} · {names[c]}", f"{size[c]}명 ({size[c]/len(df)*100:.0f}%)"] +
                [f"{prof.loc[c, f]:,} ({ratio.loc[c, f]:.2f})" for f in ["Fresh", "Milk", "Grocery", "Detergents_Paper", "Delicassen"]] +
                [f"{retail[c]}%"])
table(doc, ["Cluster (프로파일 명)", "규모", "Fresh", "Milk", "Grocery", "Deterg.", "Delica.", "Retail 비율"],
      rows, widths=[1.95, 0.75, 0.85, 0.8, 0.85, 0.85, 0.8, 0.6])
para(doc, "k=3에서는 하나로 뭉쳐 있던 'Retail형'과 'Horeca형'이 각각 규모(대형/소형)로 다시 갈라졌다. "
          "C3(27%)은 전 품목 0.94~2.36배로 매출 기여가 가장 큰 대형 종합구매 거래처(Retail 80%) — 전담 관리·볼륨 할인 대상. "
          "C0(14%)는 Grocery 1.55배·세제 1.85배지만 Fresh·Frozen은 0.2배 내외로, C3보다 작은 규모의 그로서리·세제 소매점(Retail 66%) — "
          "가공식품·생활용품 묶음 프로모션 대상으로 C3와는 별도 전략이 필요하다. "
          "C1(29%)은 Fresh 1.61배·Frozen 2.01배로 신선·냉동 위주 대형 Horeca(Retail 5%) — 신선식품 리드타임·콜드체인이 핵심 소구점. "
          "C2(30%)는 전 품목이 0.12~0.80배로 가장 소비가 작은 소형 Horeca(Retail 1%) — 매출 기여는 낮지만 이탈 방지·업셀링 타겟. "
          "k=3에서는 C0/C1(구 라벨)이 Retail 비중만으로 구분됐지만, k=4는 '같은 채널 안에서도 대형/소형 거래처가 다른 니즈를 가진다'는 "
          "정보를 추가로 드러내 실무적으로 더 실행 가능한 세분화가 되었다.",
     size=7.6, after=3)

heading(doc, "(2) 조별 토의")
table(doc, ["항목", "내용"],
      [["조별 Best Model / 이유", "[조원 결과 취합 후 기입 — 본인: K-Means k=4, log1p+StandardScaler, 6개 소비 변수 "
                                "(Silhouette 0.188, 크기 61/129/132/118). 선정 기준: 지표 단독이 아니라 ①Dendrogram 병합 거리상 "
                                "k=3·k=4가 사실상 동급이라는 구조적 근거 ②채널 내부 대형/소형 거래처까지 갈라지는 프로파일의 실행 가능성 "
                                "③클러스터 크기 균형(30/29/27/14%)]"],
       ["모델별 결과의 주요 차이", "K-Means와 Hierarchical(ward)은 최적정렬 기준 68.0% 일치 — k=3일 때(74.8%)보다 낮아져, 세분화가 "
                              "정교해질수록(k↑) centroid 기반과 분산 최소화 병합 기준의 차이가 더 크게 드러난다. "
                              "DBSCAN은 군집 수가 데이터에 의해 결정되어 eps에 극도로 민감(0.8→noise 262, 1.2→1군집)해 k를 직접 통제할 수 없었다."],
       ["Cluster별 고객 특성·비즈니스 의미", "C3 대형 종합거래처 / C1 신선·냉동 Horeca / C0 소형 그로서리·세제 소매 / C2 저지출 소형 Horeca. "
                                    "→ k=3이었다면 뭉쳐 보였을 '대형 vs 소형' 차이가 채널별로 각각 드러나 전담관리·콜드체인·묶음프로모션·"
                                    "이탈방지로 4갈래 전략 수립이 가능. [조별 논의 후 보완]"],
       ["실험 결과의 차이와 원인", "[조원 간 k(3 vs 4)·전처리(log 유무)·변수 구성 차이와 그 원인 — 논의 후 기입. "
                              "본인은 Dendrogram 절단 거리(3.98 vs 4.00)가 k=3/4 사이에 유의한 차이를 만들지 않는다는 점을 근거로 k=4를 선택함]"]],
      widths=[1.55, 5.85], fsize=7.4)

heading(doc, "(3) 개인 시사점")
para(doc, "k 선택에서 '지표가 가장 좋은 지점'과 'Dendrogram상 병합 비용이 실제로 급증하는 지점'은 다를 수 있다는 것을 체감했다. "
          "Silhouette만 보면 k=2·k=3이 k=4보다 높지만, Dendrogram 병합 거리는 k=2→3(24.45)에서 k=3→4(20.47)로 넘어갈 때와 "
          "k=3→4에서 k=4→5(16.47)로 넘어갈 때의 감소 폭이 거의 같아 k=3에서 '반드시 멈춰야 할' 구조적 근거가 없었다. 이런 경우 "
          "지표상 근소한 손해를 감수하고 한 단계 더 세분화하면, 같은 채널 안에서도 대형/소형 거래처처럼 실무적으로 더 쓸모 있는 "
          "구분이 드러날 수 있다는 것을 확인했다. 또한 log 변환 여부가 k 선택보다 결과에 더 큰 영향을 준다는 점(①)도 재확인했다 — "
          "전처리를 잘못하면 어떤 k를 골라도 한 군집에 쏠리는 것을 막을 수 없다. Clustering은 정답 라벨이 없기 때문에 Silhouette·CH "
          "같은 내부 지표는 참고 신호일 뿐이며, Dendrogram의 구조적 근거와 프로파일의 해석 가능성·비즈니스 활용도를 함께 봐야 한다는 "
          "점을 이번 k=3→4 재검토를 통해 다시 확인했다.", size=8.1)

doc.save(OUT)
print(f"Saved: {OUT}")
