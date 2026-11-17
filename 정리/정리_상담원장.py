"""상담원장_2025.xlsx 정리.

결정 기록: my-wiki/decisions/d-005-상담원장-정리-규칙.md (결정 1~5)
  결정 1  «종결» → «완료», «대기» → «보류» (CS기획팀 답)
  결정 2  처리시간 음수는 값을 두고 _표시 칸에 이유를 적는다
  결정 3  만족도 6~10은 값을 두고 _표시 칸에 이유를 적는다. 평균 만족도에서 뺀다
  결정 4  칸 이름은 원본 그대로 둔다 (상담시간은 분 단위로 바뀌어도 이름은 «상담시간»)
  결정 5  고객번호 앞뒤 공백을 없앤다
그 밖의 처리는 원본의 «작성기준» 시트와, d-005 앞서 이 정리에서 쓰던 규칙 그대로다.

원본은 읽기만 한다. 언제 돌려도 원본에서 새로 만든다 — 정리본을 다시 정리하지 않는다.
결과: 정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md

    python 정리/정리_상담원장.py
"""
import re
import sys
from pathlib import Path

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent          # cleaner/
SRC = ROOT.parent / "messy" / "상담원장_2025.xlsx"
OUT = ROOT / "정리본"
FLAG = "_표시"

CHANNEL = {"유선": "전화", "전화상담": "전화", "TEL": "전화",
           "chatbot": "챗봇", "Chatbot": "챗봇", "챗 봇": "챗봇",
           "app": "앱", "APP": "앱", "모바일앱": "앱",
           "e-mail": "이메일", "E-mail": "이메일", "메일": "이메일"}
TEAM = {"요금 팀": "요금팀", "품질 팀": "품질팀", "가입 팀": "가입팀", "해지방어 팀": "해지방어팀"}
REGION = {"경기도": "경기", "서울시": "서울", "서울특별시": "서울"}
STATUS = {"처리완료": "완료", "처리중": "진행중",
          "종결": "완료",                                # 결정 1
          "대기": "보류"}                                # 결정 1
REASK = {"y": "Y", "예": "Y", "TRUE": "Y", "n": "N", "아니오": "N", "FALSE": "N"}

NEG_TIME = "처리시간 음수 — 틀린 값이지만 진짜 값을 모름 (d-005 결정 2)"
SAT_OVER5 = "만족도 1~5점 밖 — NPS를 잘못 넣은 것으로 보이나 미확인, 평균 만족도에서 뺌 (d-005 결정 3)"

changes = {}                                           # 바꾼 것 이름 → 바뀐 원본줄 집합


def note(label, df, mask):
    changes.setdefault(label, set()).update(df.loc[mask, "_원본줄"].tolist())


def is_na(v):
    return v is None or (not isinstance(v, str) and pd.isna(v))


def text(v):
    return "" if is_na(v) else str(v)


def iso(v):
    if is_na(v):
        return v
    s = str(v).strip()
    m = re.match(r"^(\d{4})[-/.](\d{2})[-/.](\d{2})$", s)
    return f"{m[1]}-{m[2]}-{m[3]}" if m else s


def remap(df, col, table, label):
    new = df[col].map(lambda v: v if is_na(v) else table.get(str(v).strip(), str(v).strip()))
    note(label, df, df[col].map(text) != new.map(text))
    df[col] = new


def satisfaction(v):
    """1~5는 점수. 0 · - · 미응답은 빈칸(작성기준). 6~10은 값 그대로(결정 3)."""
    if is_na(v):
        return pd.NA
    s = str(v).strip().removesuffix("점")
    try:
        n = float(s)
    except ValueError:
        return pd.NA
    if n == 0:
        return pd.NA
    return int(n) if n == int(n) else n


def clean(raw):
    df = raw.copy()
    df.insert(0, "_원본줄", range(2, len(df) + 2))     # 엑셀 줄 번호(머리글 1)
    removed = []

    # 날짜 표기 하나로
    for c in ["접수일", "가입일", "최종수정일"]:
        new = df[c].map(iso)
        note(f"{c} 표기를 YYYY-MM-DD로", df, df[c].map(text) != new.map(text))
        df[c] = new

    # 표기 통일
    remap(df, "접수채널", CHANNEL, "접수채널 표기 통일 (4가지로)")
    remap(df, "상담원팀", TEAM, "상담원팀 띄어쓰기 통일")
    remap(df, "지역", REGION, "지역 표기 통일")
    remap(df, "처리상태", STATUS, "처리상태 통일 — 처리완료·종결→완료, 처리중→진행중, 대기→보류, 뒤 공백 (d-005 결정 1)")
    remap(df, "재문의", REASK, "재문의 Y/N 통일")

    # 결정 5 — 고객번호 앞뒤 공백
    new = df["고객번호"].map(lambda v: v if is_na(v) else str(v).strip())
    note("고객번호 앞뒤 공백 없앰 (d-005 결정 5)", df, df["고객번호"].map(text) != new.map(text))
    df["고객번호"] = new

    # 월정액 «89,000원» → 89000
    has_won = df["월정액"].map(text).str.contains("원")
    note("월정액 «원»·쉼표 떼기", df, has_won)
    df["월정액"] = df["월정액"].map(
        lambda v: v if is_na(v) else int(float(str(v).replace(",", "").replace("원", "").strip()))
    ).astype("Int64")

    # 만족도 — 작성기준: 미응답(0 · - · 미응답)은 빈칸
    s = df["만족도"].map(text).str.strip()
    note("만족도 0 · - · 미응답 → 빈칸 (작성기준)", df, s.isin(["0", "-", "미응답"]))
    note("만족도 «점» 떼기", df, s.str.endswith("점"))
    df["만족도"] = df["만족도"].map(satisfaction).astype("Int64")

    # 상담시간 — 값이 180을 넘으면 초 단위로 보고 분으로 바꾼다. 칸 이름은 그대로(결정 4)
    t = pd.to_numeric(df["상담시간"]).astype("Float64")
    note("상담시간 9999 → 빈칸", df, t == 9999)
    t = t.mask(t == 9999)
    sec = (t > 180).fillna(False).astype(bool)
    note("상담시간 초 → 분 (값이 180 넘는 줄)", df, sec)
    df["상담시간"] = t.mask(sec, (t / 60).round(1))

    # 결정 2 · 3 — 지우지 않고 표시
    flags = [[] for _ in range(len(df))]
    neg = (pd.to_numeric(df["처리시간"], errors="coerce") < 0).fillna(False).astype(bool)
    over5 = (pd.to_numeric(df["만족도"], errors="coerce") > 5).fillna(False).astype(bool)
    for i in range(len(df)):
        if neg.iloc[i]:
            flags[i].append(NEG_TIME)
        if over5.iloc[i]:
            flags[i].append(SAT_OVER5)
    note("_표시: 처리시간 음수 (d-005 결정 2)", df, neg)
    note("_표시: 만족도 6~10 (d-005 결정 3)", df, over5)
    df[FLAG] = [" · ".join(f) for f in flags]

    # 지우기 — 작성기준
    m = df["상담원ID"] == "AG999"
    removed.append((df[m], "테스트 접수(상담원ID AG999)"))
    df = df[~m]

    cols = [c for c in df.columns if c != "_원본줄"]
    m = df.duplicated(subset=cols, keep="first")
    removed.append((df[m], "완전 중복(날짜·표기를 맞춘 뒤 모든 칸이 같은 줄, 먼저 나온 줄을 남김)"))
    df = df[~m]

    df = df.sort_values(["접수번호", "최종수정일", "_원본줄"])
    m = df.duplicated("접수번호", keep="last")
    removed.append((df[m], "상태가 바뀌기 전 줄(같은 접수번호 중 최종수정일이 이른 줄)"))
    df = df[~m]

    df = df.sort_values(["접수일", "접수시각", "접수번호"]).reset_index(drop=True)
    return df, removed


def metrics(df):
    n = len(df)
    rq = df["재문의"]
    yn = rq.isin(["Y", "N"])
    sat = df.loc[~df[FLAG].str.contains(SAT_OVER5, regex=False), "만족도"]   # 결정 3: 만족도로 표시한 줄은 평균에서 뺌
    return {
        "상담 건수": n,
        "평균 상담시간(분)": round(float(df["상담시간"].mean()), 2),
        "재문의율(%)": round((rq == "Y").sum() / yn.sum() * 100, 2),
        "평균 만족도": round(float(sat.mean()), 2),
        "완료율(%)": round((df["처리상태"] == "완료").sum() / n * 100, 2),
        "고객 수": df["고객번호"].nunique(),
        "_재문의 분모": int(yn.sum()),
        "_만족도 응답": int(sat.notna().sum()),
    }


def main():
    raw = pd.read_excel(SRC, sheet_name="원장", dtype=object)    # 읽기만
    kept, removed = clean(raw)
    OUT.mkdir(exist_ok=True)

    kept.drop(columns="_원본줄").to_csv(OUT / "상담원장_정리.csv", index=False, encoding="utf-8-sig")

    gone = []
    for part, why in removed:
        r = raw.iloc[[i - 2 for i in part["_원본줄"]]].copy()
        r.insert(0, "원본줄", part["_원본줄"].values)
        r["지운 이유"] = why
        gone.append(r)
    gone = pd.concat(gone).sort_values("원본줄")
    gone.to_csv(OUT / "지운줄.csv", index=False, encoding="utf-8-sig")

    keep_lines = set(kept["_원본줄"])
    n_removed = sum(len(p) for p, _ in removed)
    assert len(raw) - n_removed == len(kept), "줄 수 뺄셈이 맞지 않음"
    m = metrics(kept)

    phone_sec = len(changes["상담시간 초 → 분 (값이 180 넘는 줄)"] & keep_lines)
    lines = [
        "# 상담원장 정리 — 변경 요약",
        "",
        "- 원본: `../messy/상담원장_2025.xlsx` «원장» 시트 (읽기만 함)",
        "- 스크립트: `정리/정리_상담원장.py` — 결정 기록 d-005 (결정 1~5)",
        "- 줄 수는 이 스크립트가 원본에서 새로 센 값이다.",
        "",
        "## 줄 수",
        "",
        "| 단계 | 줄 수 |",
        "|---|---|",
        f"| 원본 | {len(raw):,} |",
    ]
    lines += [f"| − {why} | −{len(p):,} |" for p, why in removed]
    lines += [f"| = 정리본 | **{len(kept):,}** |", "", "지운 줄은 `지운줄.csv`에 원본 값 그대로, 맨 끝 칸에 이유를 적었다.", ""]

    lines += [
        "## 상담시간 단위가 바뀌었다",
        "",
        f"- **`상담시간` 값이 180을 넘는 {phone_sec:,}줄을 초 단위로 보고 분으로 바꿨다** (÷60, 소수 한 자리).",
        "- 기준: 값이 180을 넘는 줄 (채널·접수일은 보지 않음).",
        "- 정리본의 `상담시간`은 모든 줄이 **분** 단위다. 칸 이름은 원본 그대로 두었다(d-005 결정 4).",
        "- 원본의 `상담시간`에는 초 단위 값이 섞여 있다. 원본과 정리본을 섞어 쓰지 않는다.",
        "",
        "## 정리본에서 바뀐 것 (정리본에 남은 줄 기준)",
        "",
        "| 바꾼 것 | 줄 수 |",
        "|---|---|",
    ]
    for label, rows in changes.items():
        lines.append(f"| {label} | {len(rows & keep_lines):,} |")
    lines += [
        "",
        f"`_표시` 칸이 채워진 줄: {(kept[FLAG] != '').sum():,}",
        "",
        "## 다섯 숫자와 고객 수",
        "",
        "| 지표 | 값 | 어떻게 셌나 |",
        "|---|---|---|",
        f"| 상담 건수 | {m['상담 건수']:,} | 접수번호 하나에 한 줄 |",
        f"| 평균 상담시간 | {m['평균 상담시간(분)']}분 | 모든 줄 분 단위 |",
        f"| 재문의율 | {m['재문의율(%)']}% | Y ÷ (Y·N) {m['_재문의 분모']:,}줄. 3/1 이전 빈칸은 분모에서 뺌 |",
        f"| 평균 만족도 | {m['평균 만족도']} | 응답 {m['_만족도 응답']:,}줄. 만족도 6~10으로 `_표시`한 줄은 뺌 (결정 3) |",
        f"| 완료율 | {m['완료율(%)']}% | `완료` ÷ 전체. 종결 포함 (결정 1) |",
        f"| 고객 수 | {m['고객 수']:,} | 서로 다른 고객번호 |",
        "",
    ]
    (OUT / "변경요약.md").write_text("\n".join(lines), encoding="utf-8")

    print(f"원본 {len(raw):,} → 정리본 {len(kept):,} (지운 줄 {n_removed:,})")
    for k, v in m.items():
        if not k.startswith("_"):
            print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
