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
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler
from scipy.cluster.hierarchy import dendrogram, linkage

OUT = Path(__file__).with_name("Clustering_실습_보고서_k3.docx")
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

km_lab = KMeans(3, random_state=42, n_init=10).fit_predict(X)
ward_lab = AgglomerativeClustering(n_clusters=3, linkage="ward").fit_predict(X)
db_lab = DBSCAN(eps=0.8, min_samples=6).fit_predict(X)

Z = linkage(X, method="ward")
merge_h = Z[:, 2]  # ascending merge heights; last k-1'th value = height forming k clusters from k+1

fig, ax = plt.subplots(figsize=(9.2, 2.3))
dendrogram(Z, truncate_mode="lastp", p=24, leaf_rotation=90, leaf_font_size=6, ax=ax,
           color_threshold=merge_h[-2])
ax.axhline(merge_h[-2], ls="--", c="crimson", lw=1)
ax.set_title("Dendrogram (ward, log+scaled) — dashed line = k=3 cut", fontsize=9)
ax.set_ylabel("Distance", fontsize=8); ax.tick_params(labelsize=6)
plt.tight_layout(); plt.savefig(SC / "fig0.png", dpi=150); plt.close()

fig, ax = plt.subplots(1, 2, figsize=(9, 2.6))
ax[0].plot(ks, inertia, "bo-"); ax[0].set_title("Elbow (log+scaled)", fontsize=9)
ax[1].plot(ks, sils, "ro-"); ax[1].set_title("Silhouette", fontsize=9)
for a in ax:
    a.set_xlabel("k", fontsize=8); a.grid(alpha=.3); a.axvline(3, ls="--", c="gray")
    a.tick_params(labelsize=7)
plt.tight_layout(); plt.savefig(SC / "fig1.png", dpi=150); plt.close()

P = PCA(2).fit_transform(X)
fig, ax = plt.subplots(1, 3, figsize=(10, 2.5))
for a, (t, lab) in zip(ax, [("K-Means k=3", km_lab), ("Hierarchical-ward k=3", ward_lab),
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

subheading(doc, "① 전처리 비교 — 지표가 좋은 모델 ≠ 쓸모 있는 세분화 (KMeans k=3)")
table(doc, ["전처리", "Silhouette", "CH Index", "클러스터 크기", "최대 클러스터 비중", "판정"],
      [["원본(raw)", "0.478", "210.2", "330 / 60 / 50", "75.0%", "금액 큰 Fresh가 지배"],
       ["Scaling만", "0.548 ★", "140.1", "393 / 45 / 2", "89.3%", "지표 1위지만 사실상 1개 군집"],
       ["log + Scaling", "0.259", "152.8", "213 / 147 / 80", "48.4%", "채택 — 균형·해석 가능"]],
      widths=[0.95, 0.8, 0.75, 1.15, 1.1, 2.4])

subheading(doc, "② 클러스터 수(k) 탐색 — Dendrogram + Elbow + Silhouette")
pic(doc, "fig0.png", 6.3)
pic(doc, "fig1.png", 5.9)
para(doc, "Dendrogram(ward) 기준 병합 거리는 k=2→3에서 24.45, k=3→4에서 20.47로 3.98 감소한다. Silhouette는 k=2(0.313)가 최고이고 "
          "k=3은 0.259. 그러나 k=2는 'Horeca vs Retail'이라는 이미 아는 Channel 정보의 재현에 가까워 마케팅 액션으로 이어지지 않는다. "
          "k=3→4 절단은 병합 거리상 k=2→3과 큰 차이가 없어(3.98 vs 4.00, 참고: k=4→5는 16.47) 더 세분화할 여지는 있으나, "
          "우선 해석이 뚜렷하고 Elbow가 완만해지는 지점과 겹치는 k=3을 기준안으로 채택 — 세분화를 더 원할 경우 k=4로 확장 가능(별도 버전 참고).",
     size=7.7, after=3)

subheading(doc, "③ Clustering Method 비교 (log+scaled, k=3)")
table(doc, ["Method", "설정", "Silhouette", "CH", "클러스터 크기", "Noise", "해석"],
      [["K-Means", "k=3, n_init=10", "0.259", "152.8", "213/147/80", "0", "★ 최종 채택"],
       ["Hierarchical", "ward, k=3", "0.255", "116.8", "262/125/53", "0", "KMeans와 74.8% 일치"],
       ["Hierarchical", "average, k=3", "0.542", "9.1", "438/1/1", "0", "singleton 분리(체이닝)"],
       ["Hierarchical", "complete, k=3", "0.298", "37.8", "402/33/5", "0", "불균형"],
       ["DBSCAN", "eps=0.8, min_s=6", "0.324", "141.3", "88/83/7", "262", "60%가 noise 처리"],
       ["DBSCAN", "eps=1.2~1.8", "-", "-", "360~418", "80→22", "전부 1개 군집으로 병합"]],
      widths=[0.85, 1.15, 0.75, 0.6, 1.0, 0.5, 2.3])
pic(doc, "fig2.png", 6.3)
para(doc, "합성 데이터 실험에서는 Moon 형태에 DBSCAN이 압도적이었으나, 실제 판매 데이터는 6차원에서 밀도가 연속적으로 이어져 "
          "eps를 조금만 키워도 전부 한 덩어리, 조금만 줄이면 60%가 noise가 되는 밀도 임계가 존재하지 않았다. 즉 데이터가 "
          "'명확히 분리된 덩어리'가 아니라 '연속 스펙트럼'일 때는 군집 수를 강제 지정하는 K-Means/Hierarchical이 적합하다. "
          "Hierarchical-average는 Silhouette 0.542로 1위지만 438/1/1로 쪼개진 결과이며, 이는 ①의 Scaling만 적용한 경우와 함께 "
          "'지표 최댓값이 곧 좋은 세분화는 아니다'를 보여준다.", size=7.8, after=3)

subheading(doc, "④ 변수 구성 비교 (KMeans k=3)")
table(doc, ["변수 구성", "Silhouette", "CH", "크기", "Channel 일치도", "해석"],
      [["6개 소비 변수", "0.259", "152.8", "213/147/80", "79.3%", "구매 패턴 기반 세분화(채택)"],
       ["8개 전체(+Channel/Region)", "0.204", "143.1", "162/145/133", "97.0%", "Channel이 군집을 지배 → 신규 정보 없음"]],
      widths=[1.7, 0.8, 0.6, 1.0, 1.0, 2.1])

subheading(doc, "⑤ Customer Profiling — K-Means k=3 (원본 스케일 평균, 괄호는 전체평균 대비 배수)")
rows = []
names = {0: "식료품·세제 특화 소매형", 1: "고액 종합구매 대형 거래처", 2: "신선식품 중심 요식업(Horeca)"}
for c in [2, 1, 0]:
    rows.append([f"C{c} · {names[c]}", f"{size[c]}명 ({size[c]/len(df)*100:.0f}%)"] +
                [f"{prof.loc[c, f]:,} ({ratio.loc[c, f]:.2f})" for f in ["Fresh", "Milk", "Grocery", "Detergents_Paper", "Delicassen"]] +
                [f"{retail[c]}%"])
table(doc, ["Cluster (프로파일 명)", "규모", "Fresh", "Milk", "Grocery", "Deterg.", "Delica.", "Retail 비율"],
      rows, widths=[1.9, 0.75, 0.85, 0.8, 0.85, 0.85, 0.8, 0.6])
para(doc, "C2(48%)는 Fresh 위주에 세제 0.15배로 Channel이 99% Horeca인 요식업 세그먼트 — 신선식품 배송 주기·품질 SLA가 핵심 소구점. "
          "C0(18%)은 Grocery 1.58배·세제 1.93배인 반면 Fresh 0.24배로, 소량 다품목 회전이 중요한 소매점형 — 가공식품·생활용품 묶음 프로모션 대상. "
          "C1(33%)은 전 품목 1.35~1.87배로 매출 기여가 가장 큰 대형 거래처 — 전담 관리·볼륨 할인·이탈 방지 대상. "
          "C0/C1은 Retail 비중이 각각 60%/62%로 유사하지만 구매 구성이 완전히 달라, Channel 라벨만으로는 얻을 수 없는 세분화가 나왔다.",
     size=7.8, after=3)

heading(doc, "(2) 조별 토의")
table(doc, ["항목", "내용"],
      [["조별 Best Model / 이유", "[조원 결과 취합 후 기입 — 본인: K-Means k=3, log1p+StandardScaler, 6개 소비 변수 "
                                "(Silhouette 0.259, 크기 213/147/80). 선정 기준: 지표 단독이 아니라 ①클러스터 크기 균형 "
                                "②프로파일 구분의 선명도 ③비즈니스 액션 도출 가능성]"],
       ["모델별 결과의 주요 차이", "K-Means와 Hierarchical(ward)은 라벨 74.8% 일치로 큰 틀은 동일, 경계 샘플에서만 차이. "
                              "반면 DBSCAN은 군집 수가 데이터에 의해 결정되어 eps에 극도로 민감(0.8→3군집+noise 262, 1.2→1군집)."],
       ["Cluster별 고객 특성·비즈니스 의미", "C2 요식업(신선식품) / C0 소매(식료품·세제) / C1 대형 종합 거래처. "
                                    "→ 각각 신선도 SLA, 묶음 프로모션, 전담 계정 관리로 채널 전략 분리 가능. [조별 논의 후 보완]"],
       ["실험 결과의 차이와 원인", "[조원 간 k·전처리(log 유무)·변수 구성 차이와 그 원인 — 논의 후 기입. "
                              "참고: Dendrogram 절단 거리상 k=3→4(20.47)과 k=2→3(24.45)의 차이(3.98)가 크지 않아, "
                              "더 세분화한 k=4 버전도 별도로 준비함(Clustering_실습_보고서.docx)]"]],
      widths=[1.55, 5.85], fsize=7.5)

heading(doc, "(3) 개인 시사점")
para(doc, "가장 큰 성능 레버는 알고리즘 선택이 아니라 전처리였다. log 변환 없이는 Silhouette가 0.548까지 올라가지만 440명 중 393명이 "
          "한 군집에 몰려 세분화로서 무가치했고, log 변환 후 지표는 0.259로 떨어졌지만 비로소 해석 가능한 3개 세그먼트가 나왔다. "
          "Clustering은 정답 라벨이 없기 때문에 Silhouette·CH 같은 내부 지표는 '모양의 좋음'만 측정할 뿐 '유용함'을 보장하지 않으며, "
          "클러스터 크기 분포와 프로파일의 해석 가능성을 반드시 함께 봐야 한다는 점을 체감했다. 또한 KNN 실습과 마찬가지로 거리 기반 "
          "기법에서는 스케일·분포 보정이 하이퍼파라미터 튜닝보다 우선한다는 공통점을 확인했다.", size=8.3)

doc.save(OUT)
print(f"Saved: {OUT}")
