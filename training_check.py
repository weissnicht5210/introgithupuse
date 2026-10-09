#!/usr/bin/env python3
"""신규 입사자 필수교육 미이수자 산출 (매월 초 실행).

사용법:
  python3 training_check.py 직원양식.xlsx 학습현황.xlsx [기준일 YYYY-MM-DD] [-o 결과.xlsx]

규칙
  - 대상: 상태가 '재직' 또는 '모자발령' 인 직원 (최종입사일 기준 대상기간 내 입사자)
  - 제외(온라인 사이트 삭제 대상): 상태 '퇴직'
  - 대상 입사기간: 기준일이 속한 달의 1일 기준 3개월 전 1일 ~ 전월 말일
      예) 기준일 2026-10-09 -> 2026-07-01 ~ 2026-09-30 입사자 (이수기한 = 입사일 + 3개월)
  - 이수: '필수과정1' 과 '필수과정2' 모두 수료일이 있어야 함
      시작일 = 두 과정 중 빠른 날짜 / 종료일, 수료일 = 두 과정 중 늦은 날짜
"""
import argparse
import re
from datetime import date, datetime

import pandas as pd

COURSES = ["2026 한양대병원 필수과정1", "2026 한양대병원 필수과정2"]
TARGET_STATUS = ["재직", "모자발령"]
EXCLUDE_STATUS = ["퇴직"]


def read_with_header(path, key):
    """key 컬럼명이 있는 행을 헤더로 자동 탐색."""
    raw = pd.read_excel(path, header=None, dtype=object)
    for i in range(min(10, len(raw))):
        if key in [str(v).strip() for v in raw.iloc[i].tolist()]:
            df = raw.iloc[i + 1:].copy()
            df.columns = [str(v).strip() for v in raw.iloc[i].tolist()]
            return df.dropna(how="all").reset_index(drop=True)
    raise SystemExit(f"{path}: '{key}' 헤더를 찾지 못했습니다.")


def to_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    m = re.match(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", s)
    return date(*map(int, m.groups())) if m else None


def add_months(d, n):
    y, m = divmod(d.year * 12 + d.month - 1 + n, 12)
    return date(y, m + 1, 1)


def deadline(d):
    """입사일 + 3개월 (말일 초과 시 말일로 보정)."""
    if d is None:
        return None
    f = add_months(d, 3)
    nxt = add_months(f, 1)
    last = (nxt - pd.Timedelta(days=1)).day
    return f.replace(day=min(d.day, last))


def norm_id(v):
    return re.sub(r"\D", "", str(v)) if v is not None else ""


def norm_phone(v):
    return re.sub(r"\D", "", str(v)) if v is not None else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("emp")
    ap.add_argument("learn")
    ap.add_argument("base", nargs="?", default=None, help="기준일(기본: 오늘)")
    ap.add_argument("-o", "--out", default="교육이수현황_결과.xlsx")
    a = ap.parse_args()

    base = to_date(a.base) if a.base else date.today()
    start = add_months(base.replace(day=1), -3)      # 대상 입사 시작일
    end = base.replace(day=1) - pd.Timedelta(days=1)  # 대상 입사 종료일(전월 말)
    end = end.date() if hasattr(end, "date") else end
    period = f"{start:%Y-%m-%d} ~ {end:%Y-%m-%d}"

    emp = read_with_header(a.emp, "사번")
    lrn = read_with_header(a.learn, "과정명")

    emp["사번"] = emp["사번"].map(norm_id)
    emp["입사일"] = emp["최종입사일"].map(to_date)
    emp["이수기한"] = emp["입사일"].map(deadline)

    # 학습현황: 사번(로그인아이디의 숫자) 우선, 없으면 연락처로 매칭
    lrn["사번"] = lrn["로그인아이디"].map(norm_id)
    lrn["폰"] = lrn["연락처"].map(norm_phone)
    lrn["시작"] = lrn["시작일"].map(to_date)
    lrn["수료"] = lrn["수료일"].map(to_date)
    lrn["종료"] = lrn["종료일"].map(to_date)

    def find_rows(r):
        m = lrn[lrn["사번"] == r["사번"]]
        if m.empty:
            m = lrn[lrn["폰"] == norm_phone(r["전화번호"])]
        return m

    rows = []
    for _, r in emp.iterrows():
        st = str(r["상태"]).strip()
        rec = {
            "사번": r["사번"], "성명": r["성명"], "소속부서": r["소속부서"],
            "근무부서": r["근무부서"], "직종": r["직종"], "전화번호": r["전화번호"],
            "상태": st, "최종입사일": r["입사일"], "이수기한": r["이수기한"],
        }
        in_period = r["입사일"] is not None and start <= r["입사일"] <= end
        if st in EXCLUDE_STATUS:
            rec["구분"] = "온라인사이트 제외(퇴직)"
        elif st in TARGET_STATUS:
            rec["구분"] = "대상" if in_period else "대상기간 외 입사"
        else:
            rec["구분"] = f"확인필요(상태:{st})"
        rec["대상입사기간"] = period

        m = find_rows(r)
        done, starts, ends, certs = [], [], [], []
        for c in COURSES:
            cm = m[m["과정명"].astype(str).str.strip() == c]
            has = (not cm.empty) and cm["수료"].notna().all()
            done.append(has)
            if not cm.empty:
                starts += [d for d in cm["시작"] if d]
                ends += [d for d in cm["종료"] if d]
                certs += [d for d in cm["수료"] if d]
        rec["필수과정1"] = "수료" if done[0] else ("수강중" if (m["과정명"] == COURSES[0]).any() else "미신청")
        rec["필수과정2"] = "수료" if done[1] else ("수강중" if (m["과정명"] == COURSES[1]).any() else "미신청")
        full = all(done)
        rec["이수여부"] = "이수" if full else "미이수"
        rec["시작일"] = min(starts) if full and starts else None
        rec["종료일"] = max(ends) if full and ends else None
        rec["수료일"] = max(certs) if full and certs else None
        rows.append(rec)

    res = pd.DataFrame(rows)
    targets = res[res["구분"] == "대상"]
    sms = targets[targets["이수여부"] == "미이수"]
    removal = res[res["구분"] == "온라인사이트 제외(퇴직)"]
    review = res[res["구분"].str.startswith("확인필요")]
    summary = pd.DataFrame({
        "항목": ["기준일", "대상 입사기간", "대상자", "이수", "미이수(문자발송)",
                 "온라인 제외(퇴직)", "상태 확인필요", "대상기간 외(재직/모자)"],
        "값": [str(base), period, len(targets), (targets["이수여부"] == "이수").sum(),
              len(sms), len(removal), len(review), (res["구분"] == "대상기간 외 입사").sum()],
    })

    with pd.ExcelWriter(a.out, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="요약", index=False)
        sms.to_excel(w, sheet_name="문자발송대상", index=False)
        targets.to_excel(w, sheet_name="대상자전체", index=False)
        removal.to_excel(w, sheet_name="온라인제외(퇴직)", index=False)
        if not review.empty:
            review.to_excel(w, sheet_name="확인필요", index=False)
        for ws in w.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = max(
                    10, min(30, max(len(str(c.value or "")) for c in col) + 2))
    print(summary.to_string(index=False))
    print(f"-> {a.out}")


if __name__ == "__main__":
    main()
