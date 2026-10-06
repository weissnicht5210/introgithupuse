"""전공의 시간외 근무 → 연속 근무 묶기 + 서울급여 해당 판정 (Microsoft 365 전용, 행 수 제한 없음)

입력: 시간외 근무 상세내역, 시간외 근무(월 합계), 서울급여 명단(전공의 파견 모니터링의 출력_서울급여를 값으로 복사)
출력: 출력_연속근무(이어진 근무를 한 줄로), 출력_서울급여판정(사람별 해당 여부·상세/합계 대조)
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

# ================================================================ 입력_시간외상세
ds = wb.create_sheet("입력_시간외상세")
ds["A1"].value, ds["A1"].font = "월별 전공의 시간외 근무 상세내역", f_title
ds["A2"].value, ds["A2"].font = "※ 원본 파일의 5행(자료 첫 행)부터 마지막 행까지 복사해 A5에 붙여넣으세요. 표가 자동으로 늘어납니다. 초록 머리글 열은 자동 계산.", f_note
ds["A3"].value, ds["A3"].font = "※ 여러 달을 이어 붙여도 됩니다(월이 바뀌어도 이어진 근무를 찾음).", f_note
d_in = ["부서", "사번", "성명", "직위", "근무일자", "시작시간", "종료시간", "06~07시", "07~18시", "18~22시", "22~06시"]
d_w = [10, 10, 8, 10, 11, 16, 16, 7, 7, 7, 7]
BANDS = ["06~07시", "07~18시", "18~22시", "22~06시"]
d_rows = (DATA or {}).get("detail_rows") or [dict(zip(d_in, r)) for r in [
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-01", "2026-09-01 06:00", "2026-09-01 20:00", 1, 11, 2, 0),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-02", "2026-09-02 06:00", "2026-09-02 20:00", 1, 11, 2, 0),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-04", "2026-09-04 06:00", "2026-09-04 20:00", 1, 11, 2, 0),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-04", "2026-09-04 20:00", "2026-09-05 08:00", 1, 1, 2, 8),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-05", "2026-09-05 20:00", "2026-09-06 08:00", 1, 1, 2, 8),
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "2026-09-24", "2026-09-24 08:00", "2026-09-24 20:00", 0, 10, 2, 0),
    ("ㅁㅁ과", 3333333, "이ㅇㅇ", "레지던트1", "2026-09-07", "2026-09-07 18:00", "2026-09-07 22:00", 0, 0, 4, 0),
    ("ㅁㅁ과", 3333333, "이ㅇㅇ", "레지던트1", "2026-09-07", "2026-09-07 07:00", "2026-09-07 18:00", 0, 11, 0, 0),
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
    ("서울급여", 8, "calc", '=IF([@사번키]="","",IF(ISNUMBER(XMATCH([@사번키],서울명단표[사번키])),"해당","비해당"))', None),
]
make_table(ds, "상세표", 4, 1, d_cols, d_rows)
ds.freeze_panes = "D5"

# ================================================================ 입력_시간외합계
ss = wb.create_sheet("입력_시간외합계")
ss["A1"].value, ss["A1"].font = "월별 전공의 시간외 근무 (합계)", f_title
ss["A2"].value, ss["A2"].font = "※ 원본 파일의 5행부터 마지막 행까지 복사해 A5에 붙여넣으세요. 시간 칸이 '20.00' 같은 글자여도 숫자로 읽습니다.", f_note
s_in = ["부서", "사번", "성명", "직위", "직위(변환)"] + BANDS
s_rows = (DATA or {}).get("summary_rows") or [dict(zip(s_in, r)) for r in [
    ("ㅇㅇ과", 2222222, "김ㅇㅇ", "레지던트2", "레지던트2", "5.00", "45.00", "12.00", "16.00"),
    ("ㅁㅁ과", 3333333, "이ㅇㅇ", "레지던트1", "레지던트1", "0.00", "11.00", "4.00", "0.00"),
    ("△△과", 4444444, "박ㅇㅇ", "레지던트3", "레지던트3", "0.00", "8.00", "0.00", "0.00"),
]]
s_cols = [(h, w, "in", None, None) for h, w in zip(s_in, [10, 10, 8, 10, 10, 8, 8, 8, 8])]
s_cols.append(("사번키", 10, "calc", '=IF(TRIM([@사번]&"")="","","S|"&TRIM([@사번]&""))', None))
make_table(ss, "합계표", 4, 1, s_cols, s_rows)
ss.freeze_panes = "D5"

# ================================================================ 입력_서울급여명단
sl = wb.create_sheet("입력_서울급여명단")
l_in = ["번호", "이름", "사번", "전공의등록번호", "수련과목", "연차", "소속병원", "구분", "현재 근무병원",
        "구리 근무 시작일", "구리 근무 종료일", "서울 근무일", "구리 근무일", "판정 근거", "점검"]
l_rows = (DATA or {}).get("seoul_rows") or [dict(zip(l_in, r)) for r in [
    (1, "김ㅇㅇ", "2222222", "11-1111", "ㅇㅇ과", "레지던트2", "서울병원", "서울 소속", "서울병원", "", "", 30, 0, "구리 근무자명단에 없음 → 소속병원 지급", "정상"),
    (2, "최ㅇㅇ", "5555555", "11-2222", "ㅇㅇ과", "레지던트1", "서울병원", "서울 소속", "서울병원", "", "", 30, 0, "구리 근무자명단에 없음 → 소속병원 지급", "정상"),
]]
l_cols = [(h, w, "in", None, "@" if h == "사번" else None) for h, w in zip(l_in, [5, 8, 10, 11, 12, 9, 9, 18, 10, 11, 11, 7, 7, 30, 14])]
l_cols.append(("사번키", 10, "calc", '=IF(TRIM([@사번]&"")="","","S|"&TRIM([@사번]&""))', None))
make_table(sl, "서울명단표", 1, 1, l_cols, l_rows)
sl.freeze_panes = "C2"

# ================================================================ 출력_연속근무
oc = wb.create_sheet("출력_연속근무")
oc["A1"].value, oc["A1"].font = "연속 근무 (이어진 시간외 근무를 한 줄로)", f_title
oc["A2"].value, oc["A2"].font = ("※ 자동으로 펼쳐지는 목록입니다(입력 금지). 같은 사번에서 앞 근무의 종료 시각과 다음 근무의 시작 시각이 이어지면(설정의 허용 공백 이내) 한 줄로 묶고, "
                                 "사번·시작 시각 순으로 빈 줄 없이 보여줍니다."), f_note
for lc, lt, vc, vf in [("A3", "연속 근무 건수", "B3", '=COUNTIF(상세표[블록시작],1)'),
                       ("C3", "2개 이상 기록을 묶은 건수", "D3", '=COUNTIFS(상세표[블록시작],1,상세표[기록 수],">1")'),
                       ("E3", "주말·휴일 포함 건수", "F3", '=COUNTIF(상세표[주말·휴일 포함],"포함")'),
                       ("G3", "36시간 초과 건수", "H3", '=COUNTIFS(상세표[블록시작],1,상세표[연속 시간],">36")'),
                       ("I3", "시각 형식 오류", "J3", '=COUNTIF(상세표[시작],"⚠*")+COUNTIF(상세표[종료],"⚠*")')]:
    oc[lc].value, oc[lc].font, oc[lc].alignment = lt, f_bold, Alignment(horizontal="right")
    oc[vc].value, oc[vc].font, oc[vc].border, oc[vc].alignment = vf, f_bold, border, center
oc.conditional_formatting.add("H3", FormulaRule(formula=["H3>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
oc.conditional_formatting.add("J3", FormulaRule(formula=["J3>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
oc_heads = [("사번", 10), ("성명", 8), ("부서", 10), ("직위", 10), ("시작", 16), ("종료", 16), ("시작 요일", 6),
            ("연속 시간", 8), ("기록 수", 6)] + [(b, 8) for b in BANDS] + [("주말·휴일 포함", 9), ("서울급여", 8)]
for i, (h, w) in enumerate(oc_heads, 1):
    head_cell(oc.cell(row=5, column=i), h)
    oc.column_dimensions[CL(i)].width = w
oc.row_dimensions[5].height = 30
fmtdt = lambda c: f'IF(ISNUMBER({c}),TEXT({c},"yyyy-mm-dd hh:mm"),{c}&"")'
oc_cols = ['상세표[사번]&""', "상세표[성명]", "상세표[부서]", "상세표[직위]", fmtdt("상세표[시작]"), fmtdt("상세표[블록종료]"),
           'IF(ISNUMBER(상세표[시작]),CHOOSE(WEEKDAY(상세표[시작]),"일","월","화","수","목","금","토"),"")',
           "상세표[연속 시간]", "상세표[기록 수]"] + [f"상세표[블록 {b}]" for b in BANDS] + ["상세표[주말·휴일 포함]", "상세표[서울급여]"]
idx = ",".join(str(i) for i in range(1, len(oc_cols) + 1))
oc_arr = "CHOOSE({" + idx + "}," + ",".join(oc_cols) + ")"
bk.spill(oc, "A6", '=IF(COUNTIF(상세표[블록시작],1)=0,"연속 근무 없음",SORT(' + filt(oc_arr, "상세표[블록시작]=1", "상세표[사번]") + ',{1,5}))',
         width=len(oc_cols))
oc.conditional_formatting.add(f"A6:{CL(len(oc_cols))}{END}", FormulaRule(formula=['AND(ISNUMBER($H6),$H6>36)'], fill=PatternFill("solid", bgColor="FFC7CE")))
oc.conditional_formatting.add(f"A6:{CL(len(oc_cols))}{END}", FormulaRule(formula=['$N6="포함"'], fill=PatternFill("solid", bgColor="FCE4D6")))
oc.conditional_formatting.add(f"I6:I{END}", FormulaRule(formula=['AND(ISNUMBER($I6),$I6>1)'], font=Font(name=FONT, bold=True, color="1F3864")))
oc.freeze_panes = "C6"

# ================================================================ 출력_서울급여판정
op = wb.create_sheet("출력_서울급여판정")
op["A1"].value, op["A1"].font = "시간외 근무자 서울급여 해당 여부", f_title
op["A2"].value, op["A2"].font = ("※ 자동 결과(입력 금지). 시간외 합계의 사번이 입력_서울급여명단에 있으면 '해당'. "
                                 "상세 합계는 입력_시간외상세의 시간을 사번별로 더한 값으로, 합계표와 다르면 빨갛게 표시합니다."), f_note
for lc, lt, vc, vf in [("A3", "시간외 인원", "B3", '=SUMPRODUCT(--(TRIM(합계표[사번]&"")<>""))'),
                       ("C3", "서울급여 해당", "D3", '=SUMPRODUCT((합계표[사번키]<>"")*ISNUMBER(XMATCH(합계표[사번키],서울명단표[사번키])))'),
                       ("E3", "비해당", "F3", "=B3-D3"),
                       ("G3", "상세·합계 불일치", "H3", f'=COUNTIF($O$6:$O${END},"불일치*")'),
                       ("I3", "상세에만 있는 사번(기록 수)", "J3", '=SUMPRODUCT((상세표[사번키]<>"")*ISNA(XMATCH(상세표[사번키],합계표[사번키])))')]:
    op[lc].value, op[lc].font, op[lc].alignment = lt, f_bold, Alignment(horizontal="right")
    op[vc].value, op[vc].font, op[vc].border, op[vc].alignment = vf, f_bold, border, center
op.conditional_formatting.add("H3", FormulaRule(formula=["H3>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
op.conditional_formatting.add("J3", FormulaRule(formula=["J3>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
op_heads = [("사번", 10), ("성명", 8), ("부서", 10), ("직위(변환)", 10)] + [(b, 8) for b in BANDS] + [
    ("합계(시간)", 9), ("서울급여", 8), ("소속병원", 9), ("구분", 16), ("판정 근거(서울급여명단)", 34), ("상세 합계", 8), ("상세 대조", 14)]
for i, (h, w) in enumerate(op_heads, 1):
    head_cell(op.cell(row=5, column=i), h)
    op.column_dimensions[CL(i)].width = w
op.row_dimensions[5].height = 30
num = lambda c: f'IFERROR(VALUE(TRIM({c}&"")),0)'
sumd = lambda b: f'SUMIFS(상세표[{b}],상세표[사번키],합계표[사번키])'
lkp = lambda col: f'XLOOKUP(합계표[사번키],서울명단표[사번키],서울명단표[{col}]&"","")'
bk.spill(op, "A6",
         '=LET(v_k,TRIM(합계표[사번]&"")<>"",v_in,ISNUMBER(XMATCH(합계표[사번키],서울명단표[사번키])),'
         f'v_1,{num("합계표[06~07시]")},v_2,{num("합계표[07~18시]")},v_3,{num("합계표[18~22시]")},v_4,{num("합계표[22~06시]")},'
         f'v_t,v_1+v_2+v_3+v_4,v_d,{sumd("06~07시")}+{sumd("07~18시")}+{sumd("18~22시")}+{sumd("22~06시")},'
         'v_ok,IF(ABS(v_t-v_d)<0.01,"일치","불일치(차이 "&TEXT(v_d-v_t,"0.00")&")"),'
         'v_a,CHOOSE({1,2,3,4,5,6,7,8,9,10,11,12,13,14,15},'
         f'합계표[사번]&"",합계표[성명],합계표[부서],합계표[직위(변환)],v_1,v_2,v_3,v_4,v_t,IF(v_in,"해당","비해당"),'
         f'{lkp("소속병원")},{lkp("구분")},{lkp("판정 근거")},v_d,v_ok),'
         f'IF(SUM(--v_k)=0,"시간외 합계 없음",{filt("v_a", "v_k", "합계표[사번]")}))', width=15)
op.conditional_formatting.add(f"A6:O{END}", FormulaRule(formula=['$J6="비해당"'], fill=PatternFill("solid", bgColor="EDEDED"), font=Font(name=FONT, color="7F7F7F")))
op.conditional_formatting.add(f"O6:O{END}", FormulaRule(formula=['LEFT($O6,3)="불일치"'], fill=PatternFill("solid", bgColor="FFC7CE"), font=Font(name=FONT, color="9C0006", bold=True)))
op.conditional_formatting.add(f"J6:J{END}", FormulaRule(formula=['$J6="해당"'], fill=PatternFill("solid", bgColor="C6E0B4"), font=Font(name=FONT, bold=True)))
op.freeze_panes = "C6"
op.column_dimensions["P"].width = 3
op["Q4"].value, op["Q4"].font = "서울급여 명단에 있으나 시간외 합계에 없는 사람", f_bold
for i, (h, w) in enumerate([("사번", 10), ("이름", 8), ("수련과목", 12), ("연차", 9)]):
    head_cell(op.cell(row=5, column=17 + i), h)
    op.column_dimensions[CL(17 + i)].width = w
bk.spill(op, "Q6", '=LET(v_m,(서울명단표[사번키]<>"")*ISNA(XMATCH(서울명단표[사번키],합계표[사번키])),'
                   'v_a,CHOOSE({1,2,3,4},서울명단표[사번]&"",서울명단표[이름],서울명단표[수련과목]&"",서울명단표[연차]&""),'
                   f'IF(SUM(v_m)=0,"없음",{filt("v_a", "v_m=1", "서울명단표[사번]")}))',
         width=4)

# ================================================================ 사용안내
gd = wb.create_sheet("사용안내", 0)
lines = [
    ("시간외 근무 — 연속 근무 묶기 · 서울급여 해당 판정 (Microsoft 365 전용, 인원·행 제한 없음)", f_title),
    ("", f_base),
    ("[시트 구분] — 탭 색: 파랑=입력, 초록=출력(자동, 입력 금지), 회색=설정", f_bold),
    ("• 입력_시간외상세: '월별 전공의 시간외 근무 상세내역'의 5행부터 끝까지를 A5에 붙여넣기 (부서·사번·성명·직위·근무일자·시작시간·종료시간·06~07·07~18·18~22·22~06시)", f_base),
    ("• 입력_시간외합계: '월별 전공의 시간외 근무'의 5행부터 끝까지를 A5에 붙여넣기 (부서·사번·성명·직위·직위(변환)·시간대 4개)", f_base),
    ("• 입력_서울급여명단: 전공의 파견 모니터링 파일의 출력_서울급여를 값으로 복사해 A2부터 붙여넣기 (번호~점검 15개 열)", f_base),
    ("• 출력_연속근무: 이어진 근무를 한 줄로 묶어 사번·시작 시각 순으로 빈 줄 없이 표시 (주말·휴일 포함, 36시간 초과, 서울급여 해당 표시)", f_base),
    ("• 출력_서울급여판정: 시간외 합계의 사람마다 서울급여 해당/비해당, 서울급여명단의 판정 근거, 상세내역과 합계의 시간 대조", f_base),
    ("• 설정: 연속으로 볼 최대 공백(분), 공휴일 목록", f_base),
    ("", f_base),
    ("[연속 근무 판단]", f_bold),
    ("• 같은 사번에서 어떤 근무의 시작 시각이 다른 근무의 종료 시각과 같거나(공백 0분) 그 사이에 있으면(겹침) 이어진 근무로 봅니다. 설정에서 허용 공백(분)을 늘릴 수 있습니다.", f_base),
    ("• 예: 09-04 06:00~20:00 + 09-04 20:00~09-05 08:00 → 09-04 06:00~09-05 08:00 (26시간, 기록 2개)", f_base),
    ("• 상세내역의 행 순서와 관계없이 시작 시각으로 찾습니다. 여러 달을 이어 붙이면 월말·월초에 걸친 근무도 이어집니다.", f_base),
    ("• 주말·휴일 포함: 연속 근무가 토·일 또는 설정의 공휴일에 하루라도 걸치면 '포함'. 36시간 초과 연속 근무는 빨간색(전공의법 연속수련 한도 참고).", f_base),
    ("", f_base),
    ("[서울급여 해당 판단]", f_bold),
    ("• 사번이 입력_서울급여명단에 있으면 '해당', 없으면 '비해당'. 사번은 숫자/글자 구분 없이 비교합니다.", f_base),
    ("• 서울급여명단은 급여 기준월이 시간외 조회년월과 같은 달의 것을 넣어야 합니다(예: 조회년월 2026-09 → 급여 기준월 2026-09).", f_base),
    ("• 상세 대조: 합계표의 시간대별 합계와 상세내역을 사번별로 더한 값이 다르면 '불일치(차이)'로 표시합니다.", f_base),
    ("", f_base),
    ("[참고]", f_bold),
    ("• Microsoft 365(또는 Excel 2021 이상)에서만 정상 동작합니다. 출력 시트에 값을 입력하면 #SPILL! 오류가 납니다.", f_base),
    ("• 예시 데이터(ㅇㅇ)는 지우고 실제 자료를 넣으세요. 개인정보가 포함되므로 파일 암호와 접근 권한 관리를 권장합니다.", f_base),
]
for i, (tx, fo) in enumerate(lines, 1):
    gd.cell(row=i, column=1, value=tx).font = fo
gd.column_dimensions["A"].width = 170
gd.sheet_view.showGridLines = False

ORDER = ["사용안내", "입력_시간외상세", "입력_시간외합계", "입력_서울급여명단", "출력_연속근무", "출력_서울급여판정", "설정"]
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
for sh in wb.worksheets:
    for pre, color in {"입력_": "2F5597", "출력_": "548235", "설정": "7F7F7F"}.items():
        if sh.title.startswith(pre):
            sh.sheet_properties.tabColor = color
bk.save(OUT)
print("saved", OUT, sum(len(v) for v in bk.spills.values()), "dynamic cells")
