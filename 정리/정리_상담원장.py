# 결정 기록: my-wiki/decisions/d-003-상담원장-정리-결정.md
#   결정 1. 처리상태 «종결»→«완료», «대기»→«보류»로 합친다
#   결정 2. 처리시간이 음수인 줄은 값을 두고 _표시 칸에 이유를 적는다
#   결정 3. 만족도 6~10점 줄은 값을 두고 _표시 칸에 이유를 적고, 평균 만족도에서는 뺀다
#   결정 4. 원본 24칸 이름은 그대로 두고, _표시 칸을 맨 끝에 더한다
#   결정 5. 고객번호 앞뒤 공백을 없앤다
# 원본(messy/상담원장_2025.xlsx)은 읽기만 한다.
import os
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                       # cleaner/
SRC = os.path.join(os.path.dirname(ROOT), "messy", "상담원장_2025.xlsx")
OUT_DIR = os.path.join(ROOT, "정리본")
OUT = os.path.join(OUT_DIR, "상담원장_정리.csv")
OUT_DEL = os.path.join(OUT_DIR, "지운줄.csv")
OUT_MD = os.path.join(OUT_DIR, "변경요약.md")

MARK_NEG = "처리시간 음수 (값은 그대로, 결정 2)"
MARK_SAT = "만족도 6~10점 (NPS를 잘못 넣은 것으로 보임, 값은 그대로, 평균에서 제외, 결정 3)"

raw = pd.read_excel(SRC, sheet_name="원장", dtype=object)
COLS = list(raw.columns)
raw["_원본줄"] = range(len(raw))
df = raw.copy()
removed = []   # (줄 DataFrame, 이유)

# 1. 테스트 접수 (AG999)
test = df["상담원ID"] == "AG999"
removed.append((df[test], "테스트 접수 (상담원ID AG999)"))
df = df[~test].copy()


# 2. 날짜 → YYYY-MM-DD
def to_date(v):
    s = str(v).strip()
    for f in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return pd.to_datetime(s, format=f)
        except ValueError:
            pass
    raise ValueError(f"날짜를 읽지 못함: {v!r}")


df["접수일"] = df["접수일"].map(to_date)
df["최종수정일"] = df["최종수정일"].map(to_date)

# 3. 접수번호 하나에 한 줄 — 최종수정일이 가장 늦은 줄, 같으면 원본에서 뒤의 줄
df = df.sort_values(["접수번호", "최종수정일", "_원본줄"])
latest = df.groupby("접수번호")["최종수정일"].transform("max")
older = df["최종수정일"] < latest
removed.append((df[older], "이전 상태 줄 (같은 접수번호에 최종수정일이 더 늦은 줄이 있음)"))
top = df[~older]
dup = top.duplicated("접수번호", keep="last")
removed.append((top[dup], "중복 추출 (같은 접수번호 · 같은 최종수정일 줄이 또 있음)"))
df = top[~dup].sort_values("_원본줄").copy()

# 4. 글자 정리
df["고객번호"] = df["고객번호"].astype(str).str.strip()          # 결정 5
df["월정액"] = df["월정액"].map(
    lambda v: v if isinstance(v, (int, float)) else int(str(v).replace(",", "").replace("원", "").strip())
)
df["지역"] = df["지역"].replace({"서울시": "서울", "서울특별시": "서울", "경기도": "경기"})

ch = {"전화": "전화", "유선": "전화", "전화상담": "전화", "TEL": "전화",
      "챗봇": "챗봇", "chatbot": "챗봇", "챗 봇": "챗봇", "Chatbot": "챗봇",
      "앱": "앱", "모바일앱": "앱", "app": "앱", "APP": "앱",
      "이메일": "이메일", "e-mail": "이메일", "메일": "이메일", "E-mail": "이메일"}
df["접수채널"] = df["접수채널"].astype(str).str.strip().map(ch)
assert df["접수채널"].notna().all()

df["상담원팀"] = df["상담원팀"].astype(str).str.replace(" ", "", regex=False)

# 결정 1: 종결→완료, 대기→보류
st = {"완료": "완료", "처리완료": "완료", "종결": "완료", "진행중": "진행중",
      "처리중": "진행중", "보류": "보류", "대기": "보류"}
df["처리상태"] = df["처리상태"].astype(str).str.strip().map(st)
assert df["처리상태"].notna().all()

yn = {"Y": "Y", "y": "Y", "예": "Y", "TRUE": "Y", "N": "N", "n": "N", "아니오": "N", "FALSE": "N"}
df["재문의"] = df["재문의"].map(lambda v: yn[str(v).strip()] if pd.notna(v) else None)


# 5. 만족도 — 0·-·미응답은 빈칸, 6~10은 값을 두고 표시(결정 3), 그 밖의 값은 멈춘다
def sat(v):
    if pd.isna(v):
        return None
    s = str(v).strip().replace("점", "")
    if s in ("0", "-", "미응답"):
        return None
    n = int(s)
    assert 1 <= n <= 10, f"만족도 범위 밖: {v!r}"
    return n


df["만족도"] = df["만족도"].map(sat).astype("Int64")

# 6. 상담시간 — 2025-07-01 이후 전화 채널은 초 → 분
df["상담시간"] = df["상담시간"].astype(float)
sec = df["상담시간"] > 180
n_sec = int(sec.sum())
df.loc[sec, "상담시간"] = (df.loc[sec, "상담시간"] / 60).round(2)

df["접수일"] = df["접수일"].dt.strftime("%Y-%m-%d")
df["최종수정일"] = df["최종수정일"].dt.strftime("%Y-%m-%d")

# _표시 칸 (결정 2 · 3)
neg = df["처리시간"].astype(float) < 0
hi = df["만족도"].notna() & (df["만족도"] > 5)
assert not (neg & hi).any()
df["_표시"] = ""
df.loc[neg, "_표시"] = MARK_NEG
df.loc[hi, "_표시"] = MARK_SAT

# ---- 쓰기 ----
os.makedirs(OUT_DIR, exist_ok=True)
clean = df[COLS + ["_표시"]]
clean.to_csv(OUT, index=False, encoding="utf-8-sig", lineterminator="\n")

parts = []
for part, why in removed:
    p = raw.loc[part["_원본줄"].values, ["_원본줄"] + COLS].copy()   # 지운 줄은 원본 값 그대로
    p["지운이유"] = why
    parts.append(p)
gone = pd.concat(parts).sort_values("_원본줄")
gone.drop(columns="_원본줄").to_csv(OUT_DEL, index=False, encoding="utf-8-sig", lineterminator="\n")

# ---- 숫자 ----
n = len(clean)
t_mean = clean["상담시간"].mean()
r = clean[clean["접수일"] >= "2025-03-01"]
resp = (r["재문의"] == "Y").sum() / r["재문의"].notna().sum() * 100
sat_ok = clean.loc[clean["_표시"] != MARK_SAT, "만족도"].dropna()   # 결정 3: 표시한 줄은 평균에서 뺀다
sat_mean = sat_ok.astype(float).mean()
done = (clean["처리상태"] == "완료").sum() / n * 100
cust = clean["고객번호"].nunique()
vc = clean["처리상태"].value_counts()

# ---- 변경요약.md: 원본 → 정리본, 칸별로 값이 바뀐 줄 수 ----
orig = raw.set_index("_원본줄").loc[df["_원본줄"].values]


def same(a, b):
    if pd.isna(a) and pd.isna(b):
        return True
    if pd.isna(a) or pd.isna(b):
        return False
    try:
        return float(a) == float(b)
    except (TypeError, ValueError):
        return str(a) == str(b)


changed = {}
for c in COLS:
    changed[c] = sum(not same(a, b) for a, b in zip(orig[c].values, clean[c].values))

ndel = {why: len(p) for p, why in removed}
lines = [
    "# 상담원장 정리 변경요약",
    "",
    "원본 `messy/상담원장_2025.xlsx` 시트 `원장`은 읽기만 했다. 결정 기록: `my-wiki/decisions/d-003-상담원장-정리-결정.md`.",
    "",
    "## 줄 수",
    f"- 원본 {len(raw):,}줄 → 정리본 {n:,}줄, 지운 줄 {len(gone):,}줄 (`지운줄.csv`)",
]
lines += [f"  - {why}: {k:,}줄" for why, k in ndel.items()]
lines += [
    "",
    "## 값이 바뀐 줄 수 (남은 줄 기준, 원본과 비교)",
    "| 칸 | 바뀐 줄 |", "|---|---|",
]
lines += [f"| {c} | {k:,} |" for c, k in changed.items() if k]
lines += [
    "",
    "- 날짜 칸은 값이 아니라 표기만 `YYYY-MM-DD`로 맞춘 줄이 있으면 위 표에 센다.",
    f"- `_표시` 칸(25번째, 새로 더함): 결정 2 {int(neg.sum())}줄 · 결정 3 {int(hi.sum())}줄 · 합 {int((neg | hi).sum())}줄. 이 줄들의 값은 바뀌지 않았다.",
    "",
    "## 상담시간 칸 — 단위가 섞여 있다",
    f"- 2025-07-01 이후 접수된 전화 줄 {n_sec:,}줄은 원본에서 **초** 단위였고, 정리본에서는 **분**(÷60, 소수 둘째 자리 반올림)으로 바꿨다.",
    "- 그 밖의 줄은 원래 분이라 그대로다. 그래서 상담시간 칸의 \"바뀐 줄\" 수는 위 환산 줄이 대부분이다.",
    "",
    "## 결정별 반영",
    f"- 결정 1: 처리상태 `완료 {vc.get('완료', 0):,} · 진행중 {vc.get('진행중', 0):,} · 보류 {vc.get('보류', 0):,}`",
    "- 결정 4: 원본 24칸 이름 그대로, `_표시` 칸만 맨 끝에 더함",
    f"- 결정 5: 고객번호 앞뒤 공백 제거. 고객 수 {cust:,}명",
    "",
    "## 숫자",
    f"- 건수 {n:,} · 상담시간 평균 {t_mean:.2f}분 · 재문의율 {resp:.2f}% · 평균 만족도 {sat_mean:.2f}(응답 {len(sat_ok):,}건, 표시한 줄 제외) · 완료율 {done:.2f}%",
    f"- 고객 수 {cust:,}명",
]
with open(OUT_MD, "w", encoding="utf-8", newline="\n") as f:
    f.write("\n".join(lines) + "\n")

print(f"건수 {n} / 상담시간 {t_mean:.2f} / 재문의율 {resp:.2f} / 만족도 {sat_mean:.2f} (응답 {len(sat_ok)}) / 완료율 {done:.2f} / 고객 {cust}")
print("지운 줄", len(gone), ndel, "| 초→분", n_sec, "| 표시", int(neg.sum()), int(hi.sum()))
