"""전공의 시간외 근무 → 연속 근무 묶기 + 서울급여 해당 판정 (Microsoft 365 전용, 행 수 제한 없음)

입력: 시간외 근무 상세내역(전월 말일 기록 포함 가능), 시간외 근무(월 합계), 기관명단(전공의 파견 모니터링의 출력_기관명단, 전월·당월·익월)
출력: 출력_연속근무, 출력_월경계당직(월을 넘는 당직의 반영 급여월·기관), 출력_서울급여판정(사람별 서울 반영 시간)
VERIFY=1 → LibreOffice 검증용, DATA=<pickle> → 실제 자료로 채우기(저장소에 넣지 않음)
"""
import os
import pickle

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill

from xlsx365 import (Book, CL, FONT, border, center, f_base, f_bold, f_head, f_input, f_note, f_title, fill_key,
                     head_cell, make_table, q, xl)

OUT = os.environ.get("OUT", "시간외근무_서울급여_판정.xlsx")
DATA = None
if os.environ.get("DATA"):
    with open(os.environ["DATA"], "rb") as fh:
        DATA = pickle.load(fh)
END = 20005   # 서식 적용 끝 행 (출력)
MAXR = 1048576
SEC = "1/86400"  # 시각 비교 여유 1초 (소수점 오차 방지)


def filt(arr, cond, rows_ref):
    """FILTER — 표에 자료가 한 행뿐일 때도 오류 없이 동작하도록 분기."""
    return f'IF(ROWS({rows_ref})=1,IF(AND({cond}),{arr},""),FILTER({arr},{cond}))'

wb = Workbook()
wb.remove(wb.active)
bk = Book(wb)

# ================================================================ 설정
st = wb.create_sheet("설정")
head_cell(st["A1"], "설정값 (노란 칸만 수정)")
head_cell(st["B1"], "값")
st["A2"].value, st["A2"].font, st["A2"].border = "연속으로 볼 최대 공백(분)", f_bold, border
c = st["B2"]
c.value, c.font, c.fill, c.border, c.alignment = (DATA or {}).get("gap_min", 0), f_input, fill_key, border, center
bk.add_name("연속허용분", f"{q('설정')}!$B$2")
st["A3"].value, st["A3"].font = "※ 0이면 앞 근무의 종료 시각과 다음 근무의 시작 시각이 같을 때만 이어진 근무로 봅니다. 예: 10 → 10분 이내 공백도 연속.", f_note
st["A4"].value, st["A4"].font, st["A4"].border = "대상월(급여 처리할 달, 1일)", f_bold, border
c = st["B4"]
c.value, c.number_format, c.font, c.fill, c.border, c.alignment = ((DATA or {}).get("target") or "=DATE(YEAR(TODAY()),MONTH(TODAY())-1,1)"), "yyyy-mm", f_input, fill_key, border, center
bk.add_name("대상월", f"{q('설정')}!$B$4")
bk.add_name("대상월키", 'TEXT(대상월,"yyyy-mm")')
st.column_dimensions["A"].width = 30
st.column_dimensions["B"].width = 12
HOLIDAYS_2026 = [
    ("2026-01-01", "신정"), ("2026-02-16", "설날 연휴"), ("2026-02-17", "설날"), ("2026-02-18", "설날 연휴"),
    ("2026-03-01", "삼일절"), ("2026-03-02", "대체공휴일(삼일절)"), ("2026-05-05", "어린이날"),
    ("2026-05-24", "부처님오신날"), ("2026-05-25", "대체공휴일(부처님오신날)"), ("2026-06-03", "전국동시지방선거"),
    ("2026-06-06", "현충일"), ("2026-08-15", "광복절"), ("2026-08-17", "대체공휴일(광복절)"),
    ("2026-09-24", "추석 연휴"), ("2026-09-25", "추석"), ("2026-09-26", "추석 연휴"), ("2026-10-03", "개천절"),
    ("2026-10-05", "대체공휴일(개천절)"), ("2026-10-09", "한글날"), ("2026-12-25", "성탄절"),
]
from datetime import date as _d
hol_rows = [{"날짜": _d.fromisoformat(d), "이름": n} for d, n in ((DATA or {}).get("holidays") or HOLIDAYS_2026)]
make_table(st, "공휴일표", 6, 1, [("날짜", 12, "in", None, "yyyy-mm-dd"), ("이름", 22, "in", None, None)], hol_rows, style="TableStyleLight1")
st["A5"].value, st["A5"].font = "공휴일 (주말·휴일 포함 여부 판단에 사용 — 표 아래에 계속 추가, 2027년은 확정 후 추가)", f_bold
bk.add_name("공휴일목록", "공휴일표[날짜]")
st["D5"].value, st["D5"].font = "구리와 오가면 '익월초 급여 반영'으로 표시할 기관", f_bold
make_table(st, "익월초기관표", 6, 4, [("기관", 14, "in", None, None)], [{"기관": h} for h in ["인천병원", "오산병원", "창원병원", "제주병원"]], style="TableStyleLight1")
st.column_dimensions["C"].width = 3
bk.add_name("익월초기관목록", "익월초기관표[기관]")

# ================================================================ 입력_시간외상세
ds = wb.create_sheet("입력_시간외상세")
ds["A1"].value, ds["A1"].font = "월별 전공의 시간외 근무 상세내역", f_title
ds["A2"].value, ds["A2"].font = "※ 원본 파일의 5행(자료 첫 행)부터 마지막 행까지 복사해 A5에 붙여넣으세요. 표가 자동으로 늘어납니다. 초록 머리글 열은 자동 계산.", f_note
ds["A3"].value, ds["A3"].font = "※ 여러 달을 이어 붙여도 됩니다(월이 바뀌어도 이어진 근무를 찾음).", f_note
d_in = ["부서", "사번", "성명", "직위", "근무일자", "시작시간", "종료시간", "06~07시", "07~18시", "18~22시", "22~06시"]
d_w = [10, 10, 8, 10, 11, 16, 16, 7, 7, 7, 7]
BANDS = ["06~07시", "07~18시", "18~22시", "22~06시"]
d_rows = (DATA or {}).get("detail_rows") or [dict(zip(d_in, r)) for r in [
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-08-31", "2026-08-31 20:00", "2026-09-01 08:00", 1, 1, 2, 8),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-01", "2026-09-01 08:00", "2026-09-01 20:00", 0, 10, 2, 0),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-04", "2026-09-04 06:00", "2026-09-04 20:00", 1, 11, 2, 0),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-04", "2026-09-04 20:00", "2026-09-05 08:00", 1, 1, 2, 8),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-24", "2026-09-24 08:00", "2026-09-24 20:00", 0, 10, 2, 0),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-30", "2026-09-30 20:00", "2026-10-01 08:00", 1, 1, 2, 8),
    ("수련교육부", 3333333, "이ㅇㅇ", "인턴", "2026-08-31", "2026-08-31 20:00", "2026-09-01 08:00", 1, 1, 2, 8),
    ("수련교육부", 3333333, "이ㅇㅇ", "인턴", "2026-09-07", "2026-09-07 18:00", "2026-09-07 22:00", 0, 0, 4, 0),
    ("수련교육부", 3333333, "이ㅇㅇ", "인턴", "2026-09-07", "2026-09-07 07:00", "2026-09-07 18:00", 0, 11, 0, 0),
    ("수련교육부", 4444444, "박ㅇㅇ", "인턴", "2026-09-30", "2026-09-30 20:00", "2026-10-01 08:00", 1, 1, 2, 8),
]]


def ptime(col):
    return (f'=LET(v_x,[@{col}],IF(ISNUMBER(v_x),v_x,IF(TRIM(v_x&"")="","",IFERROR(VALUE(TRIM(v_x)),'
            f'IFERROR(DATEVALUE(LEFT(TRIM(v_x),10))+VALUE(MID(TRIM(v_x),12,2))/24+VALUE(MID(TRIM(v_x),15,2))/1440,"⚠"&v_x)))))')


NEXT = 'IF([@다음블록시작]="",1E+9,[@다음블록시작])'
KEYCOND = '상세표[사번키],[@사번키]'
RANGE = f'상세표[시작],">="&([@시작]-{SEC}),상세표[시작],"<"&({NEXT}-{SEC})'
d_cols = [(h, w, "in", None, "0.00" if h in BANDS else None) for h, w in zip(d_in, d_w)]
d_cols += [
    ("사번키", 10, "calc", '=IF(TRIM([@사번]&"")="","","S|"&TRIM([@사번]&""))', None),
    ("시작", 15, "calc", ptime("시작시간"), "yyyy-mm-dd hh:mm"),
    ("종료", 15, "calc", ptime("종료시간"), "yyyy-mm-dd hh:mm"),
    ("이어짐", 7, "calc",
     f'=IF(NOT(ISNUMBER([@시작])),"",IF(COUNTIFS({KEYCOND},상세표[시작],"<"&([@시작]-{SEC}),'
     f'상세표[종료],">="&([@시작]-연속허용분/1440-{SEC}))>0,"이어짐",""))', None),
    ("블록시작", 7, "calc", '=IF(AND(ISNUMBER([@시작]),ISNUMBER([@종료]),[@이어짐]=""),1,0)', "0"),
    ("다음블록시작", 15, "calc",
     f'=IF([@블록시작]<>1,"",LET(v_n,MINIFS(상세표[시작],{KEYCOND},상세표[블록시작],1,상세표[시작],">"&([@시작]+{SEC})),IF(v_n=0,"",v_n)))',
     "yyyy-mm-dd hh:mm"),
    ("블록종료", 15, "calc", f'=IF([@블록시작]<>1,"",MAXIFS(상세표[종료],{KEYCOND},{RANGE}))', "yyyy-mm-dd hh:mm"),
    ("연속 시간", 8, "calc", '=IF([@블록시작]<>1,"",ROUND(([@블록종료]-[@시작])*24,2))', "0.00"),
    ("기록 수", 6, "calc", f'=IF([@블록시작]<>1,"",COUNTIFS({KEYCOND},{RANGE}))', "0"),
]
for b in BANDS:
    d_cols.append((f"블록 {b}", 8, "calc", f'=IF([@블록시작]<>1,"",SUMIFS(상세표[{b}],{KEYCOND},{RANGE}))', "0.00"))
d_cols += [
    ("주말·휴일 포함", 9, "calc",
     '=IF([@블록시작]<>1,"",LET(v_a,INT([@시작]),v_b,INT([@블록종료]-' + SEC + '),'
     'IF((v_b-v_a+1)-NETWORKDAYS(v_a,v_b,공휴일목록)>0,"포함","")))', None),
    ("시작월", 8, "calc", '=IF([@블록시작]<>1,"",TEXT([@시작],"yyyy-mm"))', None),
    ("블록종료월", 8, "calc", f'=IF([@블록시작]<>1,"",TEXT([@블록종료]-{SEC},"yyyy-mm"))', None),
    ("월 넘김 포함", 8, "calc", f'=IF([@블록시작]<>1,"",IF(COUNTIFS({KEYCOND},{RANGE},상세표[월 넘김],"월 넘김")>0,"포함",""))', None),
    # ---- 기록(당직 1건) 단위: 월을 넘는 기록의 반영 급여월·기관
    ("기록 시간외", 8, "calc", '=N([@06~07시])+N([@07~18시])+N([@18~22시])+N([@22~06시])', "0.00"),
    ("기록월", 8, "calc", '=IF(ISNUMBER([@시작]),TEXT([@시작],"yyyy-mm"),"")', None),
    ("종료월", 8, "calc", f'=IF(ISNUMBER([@종료]),TEXT([@종료]-{SEC},"yyyy-mm"),"")', None),
    ("월 넘김", 7, "calc", '=IF(AND([@기록월]<>"",[@종료월]<>"",[@기록월]<>[@종료월]),"월 넘김","")', None),
    ("시작월 기관", 10, "calc", '=IF([@기록월]="","",XLOOKUP([@사번키]&"|"&[@기록월],기관표[키],기관표[급여지급병원(기준월)]&"",""))', None),
    ("종료월 기관", 10, "calc",
     '=IF([@월 넘김]="",[@시작월 기관],XLOOKUP([@사번키]&"|"&[@종료월],기관표[키],기관표[급여지급병원(기준월)]&"",""))', None),
    ("월경계 판정", 24, "calc",
     '=IF([@월 넘김]="","",'
     'IF(OR([@시작월 기관]="",[@종료월 기관]=""),"기관 정보 없음(해당 월 기관명단 필요)",'
     'IF(OR(LEFT([@시작월 기관],1)="⚠",LEFT([@종료월 기관],1)="⚠"),"기관 판정 불가(15일 판정불가)",'
     'IF([@시작월 기관]=[@종료월 기관],"같은 기관",'
     'IF(AND(OR([@시작월 기관]="서울병원",[@시작월 기관]="구리병원"),OR([@종료월 기관]="서울병원",[@종료월 기관]="구리병원")),"서울↔구리: 15일 이상 근무 기관",'
     'IF(OR(AND([@시작월 기관]="구리병원",COUNTIF(익월초기관목록,[@종료월 기관])>0),AND([@종료월 기관]="구리병원",COUNTIF(익월초기관목록,[@시작월 기관])>0)),"익월초 급여 반영",'
     '"규정 확인 필요"))))))', None),
    ("반영월", 8, "calc", '=IF([@기록월]="","",IF([@월경계 판정]="익월초 급여 반영",[@종료월],[@기록월]))', None),
    ("반영 기관", 22, "calc",
     '=IF([@기록월]="","",IF([@월경계 판정]="익월초 급여 반영","확인 필요("&[@시작월 기관]&"→"&[@종료월 기관]&")",[@시작월 기관]))', None),
]
make_table(ds, "상세표", 4, 1, d_cols, d_rows)
ds.freeze_panes = "D5"

# ================================================================ 입력_시간외합계
ss = wb.create_sheet("입력_시간외합계")
ss["A1"].value, ss["A1"].font = "월별 전공의 시간외 근무 (합계)", f_title
ss["A2"].value, ss["A2"].font = "※ 원본 파일의 5행부터 마지막 행까지 복사해 A5에 붙여넣으세요. 시간 칸이 '20.00' 같은 글자여도 숫자로 읽습니다.", f_note
s_in = ["부서", "사번", "성명", "직위", "직위(변환)"] + BANDS
s_rows = (DATA or {}).get("summary_rows") or [dict(zip(s_in, r)) for r in [
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "레지던트2", "3.00", "33.00", "10.00", "16.00"),
    ("수련교육부", 3333333, "이ㅇㅇ", "인턴", "인턴", "0.00", "11.00", "4.00", "0.00"),
    ("수련교육부", 4444444, "박ㅇㅇ", "인턴", "인턴", "1.00", "1.00", "2.00", "8.00"),
]]
s_cols = [(h, w, "in", None, None) for h, w in zip(s_in, [10, 10, 8, 10, 10, 8, 8, 8, 8])]
s_cols.append(("사번키", 10, "calc", '=IF(TRIM([@사번]&"")="","","S|"&TRIM([@사번]&""))', None))
make_table(ss, "합계표", 4, 1, s_cols, s_rows)
ss.freeze_panes = "D5"

# ================================================================ 입력_기관명단 (월별 근무·급여 기관)
sl = wb.create_sheet("입력_기관명단")
l_in = ["기준월", "사번", "이름", "수련과목", "연차", "소속병원", "기준월 근무병원", "급여지급병원(기준월)", "급여 판정 근거"]
l_rows = (DATA or {}).get("org_rows") or [dict(zip(l_in, r)) for r in [
    ("2026-08", "2222222", "김ㅇㅇ", "ㅇㅇ과", "레지던트2", "서울병원", "서울병원", "서울병원", "예시"),
    ("2026-09", "2222222", "김ㅇㅇ", "ㅇㅇ과", "레지던트2", "서울병원", "서울병원", "서울병원", "예시"),
    ("2026-10", "2222222", "김ㅇㅇ", "ㅇㅇ과", "레지던트2", "서울병원", "서울병원", "구리병원", "예시"),
    ("2026-08", "3333333", "이ㅇㅇ", "인턴", "인턴", "서울병원", "제주병원", "제주병원", "예시"),
    ("2026-09", "3333333", "이ㅇㅇ", "인턴", "인턴", "서울병원", "구리병원", "구리병원", "예시"),
    ("2026-10", "3333333", "이ㅇㅇ", "인턴", "인턴", "서울병원", "구리병원", "구리병원", "예시"),
    ("2026-08", "4444444", "박ㅇㅇ", "인턴", "인턴", "구리병원", "구리병원", "구리병원", "예시"),
    ("2026-09", "4444444", "박ㅇㅇ", "인턴", "인턴", "구리병원", "구리병원", "구리병원", "예시"),
    ("2026-10", "4444444", "박ㅇㅇ", "인턴", "인턴", "구리병원", "인천병원", "인천병원", "예시"),
]]
l_cols = [(h, w, "in", None, "@" if h in ("사번", "기준월") else None) for h, w in zip(l_in, [9, 10, 8, 12, 9, 10, 11, 13, 34])]
l_cols += [
    ("월키", 8, "calc", '=IF(ISNUMBER([@기준월]),TEXT([@기준월],"yyyy-mm"),LEFT(TRIM([@기준월]&""),7))', None),
    ("사번키", 10, "calc", '=IF(TRIM([@사번]&"")="","","S|"&TRIM([@사번]&""))', None),
    ("키", 16, "calc", '=IF([@사번키]="","",[@사번키]&"|"&[@월키])', None),
]
make_table(sl, "기관표", 1, 1, l_cols, l_rows)
sl.freeze_panes = "C2"

# ================================================================ 출력_연속근무
oc = wb.create_sheet("출력_연속근무")
oc["A1"].value, oc["A1"].font = "연속 근무 (이어진 시간외 근무를 한 줄로)", f_title
oc["A2"].value, oc["A2"].font = ("※ 자동 목록(입력 금지). 대상월에 걸친 연속 근무만 사번·시작 시각 순으로 빈 줄 없이 보여줍니다. 전월 말일에 시작한 당직도 포함됩니다. "
                                 "월 넘김 표시가 있으면 출력_월경계당직에서 반영 급여월·기관을 확인하세요."), f_note
TOUCH = '((상세표[블록시작]=1)*((상세표[시작월]=대상월키)+(상세표[블록종료월]=대상월키))>0)'
for lc, lt, vc, vf in [("A3", "연속 근무 건수", "B3", f'=SUMPRODUCT(--{TOUCH})'),
                       ("C3", "2개 이상 기록을 묶은 건수", "D3", f'=SUMPRODUCT({TOUCH}*(IFERROR(--상세표[기록 수],0)>1))'),
                       ("E3", "주말·휴일 포함", "F3", f'=SUMPRODUCT({TOUCH}*(상세표[주말·휴일 포함]="포함"))'),
                       ("G3", "36시간 초과", "H3", f'=SUMPRODUCT({TOUCH}*(IFERROR(--상세표[연속 시간],0)>36))'),
                       ("I3", "시각 형식 오류", "J3", '=COUNTIF(상세표[시작],"⚠*")+COUNTIF(상세표[종료],"⚠*")')]:
    oc[lc].value, oc[lc].font, oc[lc].alignment = lt, f_bold, Alignment(horizontal="right")
    oc[vc].value, oc[vc].font, oc[vc].border, oc[vc].alignment = vf, f_bold, border, center
oc.conditional_formatting.add("H3", FormulaRule(formula=["H3>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
oc.conditional_formatting.add("J3", FormulaRule(formula=["J3>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
oc_heads = [("사번", 10), ("성명", 8), ("부서", 10), ("직위", 10), ("시작", 16), ("종료", 16), ("시작 요일", 6),
            ("연속 시간", 8), ("기록 수", 6)] + [(b, 8) for b in BANDS] + [("주말·휴일 포함", 9), ("월 넘김 포함", 9)]
for i, (h, w) in enumerate(oc_heads, 1):
    head_cell(oc.cell(row=5, column=i), h)
    oc.column_dimensions[CL(i)].width = w
oc.row_dimensions[5].height = 30
fmtdt = lambda c: f'IF(ISNUMBER({c}),TEXT({c},"yyyy-mm-dd hh:mm"),{c}&"")'
oc_cols = ['상세표[사번]&""', "상세표[성명]", "상세표[부서]", "상세표[직위]", fmtdt("상세표[시작]"), fmtdt("상세표[블록종료]"),
           'IF(ISNUMBER(상세표[시작]),CHOOSE(WEEKDAY(상세표[시작]),"일","월","화","수","목","금","토"),"")',
           "상세표[연속 시간]", "상세표[기록 수]"] + [f"상세표[블록 {b}]" for b in BANDS] + ["상세표[주말·휴일 포함]", "상세표[월 넘김 포함]"]
idx = ",".join(str(i) for i in range(1, len(oc_cols) + 1))
oc_arr = "CHOOSE({" + idx + "}," + ",".join(oc_cols) + ")"
bk.spill(oc, "A6", f'=IF(SUMPRODUCT(--{TOUCH})=0,"연속 근무 없음",SORT(' + filt(oc_arr, TOUCH, "상세표[사번]") + ',{1,5}))',
         width=len(oc_cols))
oc.conditional_formatting.add(f"A6:{CL(len(oc_cols))}{END}", FormulaRule(formula=['AND(ISNUMBER($H6),$H6>36)'], fill=PatternFill("solid", bgColor="FFC7CE")))
oc.conditional_formatting.add(f"A6:{CL(len(oc_cols))}{END}", FormulaRule(formula=['$O6="포함"'], fill=PatternFill("solid", bgColor="DDEBF7")))
oc.conditional_formatting.add(f"A6:{CL(len(oc_cols))}{END}", FormulaRule(formula=['$N6="포함"'], fill=PatternFill("solid", bgColor="FCE4D6")))
oc.freeze_panes = "C6"

# ================================================================ 출력_월경계당직
ob = wb.create_sheet("출력_월경계당직")
ob["A1"].value, ob["A1"].font = "월을 넘는 당직 (전월 말일→이번 달, 이번 달 말일→익월)", f_title
ob["A2"].value, ob["A2"].font = ("※ 같은 기관이면 시작한 달에 반영. 서울↔구리는 시작한 달의 15일 이상 근무 기관(그 달 급여지급병원)에 반영. "
                                 "구리↔인천·오산·창원·제주는 '익월초 급여 반영'(종료한 달 급여). 그 밖의 조합은 '규정 확인 필요'."), f_note
CROSS = '(상세표[월 넘김]="월 넘김")*((상세표[기록월]=대상월키)+(상세표[종료월]=대상월키)>0)'
for lc, lt, vc, vf in [("A3", "월 넘김 당직", "B3", f'=SUMPRODUCT({CROSS})'),
                       ("C3", "전월→이번 달", "D3", f'=SUMPRODUCT({CROSS}*(상세표[종료월]=대상월키))'),
                       ("E3", "이번 달→익월", "F3", f'=SUMPRODUCT({CROSS}*(상세표[기록월]=대상월키))'),
                       ("G3", "익월초 급여 반영", "H3", f'=SUMPRODUCT({CROSS}*(상세표[월경계 판정]="익월초 급여 반영"))'),
                       ("I3", "기관 정보 없음·확인 필요", "J3",
                        f'=SUMPRODUCT({CROSS}*((LEFT(상세표[월경계 판정],2)="기관")+(상세표[월경계 판정]="규정 확인 필요")>0))')]:
    ob[lc].value, ob[lc].font, ob[lc].alignment = lt, f_bold, Alignment(horizontal="right")
    ob[vc].value, ob[vc].font, ob[vc].border, ob[vc].alignment = vf, f_bold, border, center
ob.conditional_formatting.add("J3", FormulaRule(formula=["J3>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
ob_heads = [("사번", 10), ("성명", 8), ("구분", 12), ("시작", 16), ("종료", 16), ("시간외 시간", 8),
            ("시작월 기관", 11), ("종료월 기관", 11), ("판정", 26), ("반영 급여월", 9), ("반영 기관", 26)]
for i, (h, w) in enumerate(ob_heads, 1):
    head_cell(ob.cell(row=5, column=i), h)
    ob.column_dimensions[CL(i)].width = w
ob.row_dimensions[5].height = 30
ob_cols = ['상세표[사번]&""', "상세표[성명]", 'IF(상세표[종료월]=대상월키,"전월→이번 달","이번 달→익월")',
           fmtdt("상세표[시작]"), fmtdt("상세표[종료]"), "상세표[기록 시간외]", "상세표[시작월 기관]", "상세표[종료월 기관]",
           "상세표[월경계 판정]", "상세표[반영월]", "상세표[반영 기관]"]
ob_arr = "CHOOSE({" + ",".join(str(i) for i in range(1, len(ob_cols) + 1)) + "}," + ",".join(ob_cols) + ")"
bk.spill(ob, "A6", f'=IF(SUMPRODUCT({CROSS})=0,"월을 넘는 당직 없음",SORT(' + filt(ob_arr, f"{CROSS}>0", "상세표[사번]") + ',{1,4}))',
         width=len(ob_cols))
ob.conditional_formatting.add(f"A6:K{END}", FormulaRule(formula=['$I6="익월초 급여 반영"'], fill=PatternFill("solid", bgColor="FFF2CC")))
ob.conditional_formatting.add(f"A6:K{END}", FormulaRule(formula=['OR(LEFT($I6,2)="기관",$I6="규정 확인 필요")'], fill=PatternFill("solid", bgColor="FFC7CE")))
ob.freeze_panes = "C6"

# ================================================================ 출력_서울급여판정
op = wb.create_sheet("출력_서울급여판정")
op["A1"].value, op["A1"].font = "시간외 근무 서울급여 판정 (대상월)", f_title
op["A2"].value, op["A2"].font = ("※ 자동 결과(입력 금지). 서울급여 = 입력_기관명단에서 대상월 급여지급병원이 서울병원. "
                                 "대상월 반영 시간 = 상세 기록 중 반영 급여월이 대상월인 시간(이월 −, 전입 + 반영). 서울 반영 시간 = 그중 반영 기관이 서울병원인 시간."), f_note
for lc, lt, vc, vf in [("A3", "대상월", "B3", "=대상월키"),
                       ("C3", "시간외 인원", "D3", '=SUMPRODUCT(--(TRIM(합계표[사번]&"")<>""))'),
                       ("E3", "서울급여 해당", "F3", f'=COUNTIF($L$6:$L${END},"해당")'),
                       ("G3", "상세·합계 불일치", "H3", f'=COUNTIF($G$6:$G${END},"불일치*")'),
                       ("I3", "기관명단에 없음", "J3", f'=COUNTIF($L$6:$L${END},"기관 정보 없음")')]:
    op[lc].value, op[lc].font, op[lc].alignment = lt, f_bold, Alignment(horizontal="right")
    op[vc].value, op[vc].font, op[vc].border, op[vc].alignment = vf, f_bold, border, center
op.conditional_formatting.add("H3", FormulaRule(formula=["H3>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
op.conditional_formatting.add("J3", FormulaRule(formula=["J3>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
op_heads = [("사번", 10), ("성명", 8), ("부서", 10), ("직위(변환)", 10), ("시스템 합계", 8), ("상세 합계(대상월 기록)", 9),
            ("상세 대조", 13), ("이월(−) 익월초", 8), ("전입(+) 전월 말일", 8), ("대상월 반영 시간", 9),
            ("대상월 급여지급병원", 11), ("서울급여", 9)] + [(f"서울 {b}", 8) for b in BANDS] + [("서울 반영 합계", 9)]
for i, (h, w) in enumerate(op_heads, 1):
    head_cell(op.cell(row=5, column=i), h)
    op.column_dimensions[CL(i)].width = w
op.row_dimensions[5].height = 42
num = lambda c: f'IFERROR(VALUE(TRIM({c}&"")),0)'
K = "합계표[사번키]"
sm = lambda col, *crit: f'SUMIFS(상세표[{col}],상세표[사번키],{K}' + "".join("," + c for c in crit) + ")"
C_REC_M, C_REF_M = "상세표[기록월],대상월키", "상세표[반영월],대상월키"
C_REC_NOT, C_REF_NOT = '상세표[기록월],"<>"&대상월키', '상세표[반영월],"<>"&대상월키'
C_SEOUL = '상세표[반영 기관],"서울병원"'
SEOUL_BANDS = "".join(f"v_b{i}," + sm(b, C_REF_M, C_SEOUL) + "," for i, b in enumerate(BANDS, 1))
bk.spill(op, "A6",
         '=LET(v_k,TRIM(합계표[사번]&"")<>"",'
         f'v_t,{num("합계표[06~07시]")}+{num("합계표[07~18시]")}+{num("합계표[18~22시]")}+{num("합계표[22~06시]")},'
         f'v_d,{sm("기록 시간외", C_REC_M)},'
         'v_ok,IF(ABS(v_t-v_d)<0.01,"일치","불일치(차이 "&TEXT(v_d-v_t,"0.00")&")"),'
         f'v_out,{sm("기록 시간외", C_REC_M, C_REF_NOT)},v_in,{sm("기록 시간외", C_REC_NOT, C_REF_M)},'
         f'v_r,{sm("기록 시간외", C_REF_M)},'
         'v_h,XLOOKUP(합계표[사번키]&"|"&대상월키,기관표[키],기관표[급여지급병원(기준월)]&"","기관 정보 없음"),'
         'v_s,IF(v_h="기관 정보 없음",v_h,IF(v_h="서울병원","해당","비해당")),'
         + SEOUL_BANDS +
         'v_a,CHOOSE({1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17},합계표[사번]&"",합계표[성명],합계표[부서],합계표[직위(변환)],'
         'v_t,v_d,v_ok,v_out,v_in,v_r,v_h,v_s,v_b1,v_b2,v_b3,v_b4,v_b1+v_b2+v_b3+v_b4),'
         f'IF(SUM(--v_k)=0,"시간외 합계 없음",{filt("v_a", "v_k", "합계표[사번]")}))', width=17)
op.conditional_formatting.add(f"A6:Q{END}", FormulaRule(formula=['$L6="비해당"'], fill=PatternFill("solid", bgColor="EDEDED"), font=Font(name=FONT, color="7F7F7F")))
op.conditional_formatting.add(f"L6:L{END}", FormulaRule(formula=['$L6="기관 정보 없음"'], fill=PatternFill("solid", bgColor="FFC7CE")))
op.conditional_formatting.add(f"G6:G{END}", FormulaRule(formula=['LEFT($G6,3)="불일치"'], fill=PatternFill("solid", bgColor="FFC7CE"), font=Font(name=FONT, color="9C0006", bold=True)))
op.conditional_formatting.add(f"H6:I{END}", FormulaRule(formula=['AND(ISNUMBER(H6),H6<>0)'], fill=PatternFill("solid", bgColor="FFF2CC"), font=Font(name=FONT, bold=True)))
op.conditional_formatting.add(f"L6:L{END}", FormulaRule(formula=['$L6="해당"'], fill=PatternFill("solid", bgColor="C6E0B4"), font=Font(name=FONT, bold=True)))
op.freeze_panes = "C6"
op.column_dimensions["R"].width = 3
op["S4"].value, op["S4"].font = "대상월 기록이 합계표에 없는 사번 (전월 말일 당직만 있는 사람 등)", f_bold
for i, (h, w) in enumerate([("사번", 10), ("성명", 8), ("대상월 반영 시간", 10)]):
    head_cell(op.cell(row=5, column=19 + i), h)
    op.column_dimensions[CL(19 + i)].width = w
bk.spill(op, "S6", '=LET(v_m,(상세표[반영월]=대상월키)*ISNA(XMATCH(상세표[사번키],합계표[사번키])),'
                   'v_a,CHOOSE({1,2,3},상세표[사번]&"",상세표[성명],상세표[기록 시간외]),'
                   f'IF(SUM(v_m)=0,"없음",{filt("v_a", "v_m=1", "상세표[사번]")}))', width=3)

# ================================================================ 사용안내
gd = wb.create_sheet("사용안내", 0)
lines = [
    ("시간외 근무 — 연속 근무 · 월 넘김 당직 · 서울급여 판정 (Microsoft 365 전용, 인원·행 제한 없음)", f_title),
    ("", f_base),
    ("[매달 하는 일]", f_bold),
    ("1) 설정 B4에 대상월(급여 처리할 달) 입력 — 기본값은 지난달", f_base),
    ("2) 입력_시간외상세: 대상월 상세내역 + 전월 상세내역 중 말일 기록(또는 전월 전체)을 A5부터 붙여넣기. 익월 상세가 이미 있으면 함께 넣어도 됩니다.", f_base),
    ("3) 입력_시간외합계: 대상월 '월별 전공의 시간외 근무'의 5행부터 끝까지를 A5에 붙여넣기", f_base),
    ("4) 입력_기관명단: 전공의 파견 모니터링 파일의 출력_기관명단을 값으로 복사해 붙여넣기 — 전월·대상월·익월 3개월치 (모니터링 파일 설정의 급여 기준월을 바꿔 가며 복사)", f_base),
    ("", f_base),
    ("[출력 시트]", f_bold),
    ("• 출력_연속근무: 대상월에 걸친 이어진 근무를 한 줄로(전월 말일에 시작한 당직 포함), 사번·시작 순, 빈 줄 없음. 주말·휴일, 36시간 초과, 월 넘김 표시", f_base),
    ("• 출력_월경계당직: 전월 말일→대상월, 대상월 말일→익월로 넘어가는 당직마다 시작월·종료월 기관, 판정, 반영 급여월·기관", f_base),
    ("• 출력_서울급여판정: 사람별 시스템 합계·상세 대조, 이월(−)/전입(+), 대상월 반영 시간, 대상월 급여지급병원, 서울 반영 시간(시간대별)", f_base),
    ("", f_base),
    ("[월을 넘는 당직 규칙] — 기록(당직 1건) 단위로 판단", f_bold),
    ("• 시작한 달과 끝난 달의 기관(입력_기관명단의 급여지급병원)이 같으면: 시작한 달 급여에 반영", f_base),
    ("• 서울↔구리: 시작한 달의 15일 이상 근무 기관(그 달 급여지급병원)에 반영", f_base),
    ("• 구리↔인천·오산·창원·제주(설정의 '익월초 기관' 표): '익월초 급여 반영' — 끝난 달(익월) 급여로 넘기고 반영 기관은 '확인 필요'로 표시", f_base),
    ("• 그 밖의 조합(예: 서울↔인천)은 '규정 확인 필요', 기관명단에 해당 월이 없으면 '기관 정보 없음'으로 표시", f_base),
    ("• 예: 08-31 20:00~09-01 08:00 제주→구리 → 9월(익월초) 급여 반영 / 09-30 20:00~10-01 08:00 서울→구리 → 9월 15일 이상 근무 기관(9월 급여지급병원)", f_base),
    ("", f_base),
    ("[연속 근무 판단]", f_bold),
    ("• 같은 사번에서 시작 시각이 다른 근무의 종료 시각과 같거나 겹치면 이어진 근무입니다(설정의 허용 공백(분) 이내 포함). 행 순서와 관계없이 시작 시각으로 찾습니다.", f_base),
    ("• 주말·휴일 포함: 토·일 또는 설정의 공휴일에 하루라도 걸치면 '포함'. 36시간 초과 연속 근무는 빨간색(전공의법 연속수련 한도 참고).", f_base),
    ("", f_base),
    ("[참고]", f_bold),
    ("• 시스템 합계와 비교하는 '상세 합계'는 시작 시각이 대상월인 기록의 합입니다(전월 말일 기록은 제외).", f_base),
    ("• Microsoft 365(또는 Excel 2021 이상)에서만 정상 동작합니다. 출력 시트에 값을 입력하면 #SPILL! 오류가 납니다.", f_base),
    ("• 예시 데이터(ㅇㅇ)는 지우고 실제 자료를 넣으세요. 개인정보가 포함되므로 파일 암호와 접근 권한 관리를 권장합니다.", f_base),
]
for i, (tx, fo) in enumerate(lines, 1):
    gd.cell(row=i, column=1, value=tx).font = fo
gd.column_dimensions["A"].width = 170
gd.sheet_view.showGridLines = False

ORDER = ["사용안내", "입력_시간외상세", "입력_시간외합계", "입력_기관명단", "출력_연속근무", "출력_월경계당직", "출력_서울급여판정", "설정"]
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
for sh in wb.worksheets:
    for pre, color in {"입력_": "2F5597", "출력_": "548235", "설정": "7F7F7F"}.items():
        if sh.title.startswith(pre):
            sh.sheet_properties.tabColor = color
bk.save(OUT)
print("saved", OUT, sum(len(v) for v in bk.spills.values()), "dynamic cells")
