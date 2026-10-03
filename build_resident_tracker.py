from datetime import date

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = "전공의_파견_모니터링.xlsx"
FONT = "맑은 고딕"
CAPACITY = 200   # 전공의명단(파견 관리 대상) 행 수
BASE_CAP = 500   # 기본자료 붙여넣기 가능 행 수
LAST = CAPACITY + 1
BLAST = BASE_CAP + 1

f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_input = Font(name=FONT, size=10, color="0000FF")  # 입력값=파랑, 수식=검정
f_title = Font(name=FONT, size=14, bold=True)
fill_head = PatternFill("solid", fgColor="1F3864")
fill_head2 = PatternFill("solid", fgColor="548235")  # 기본자료 자동조회 열 머리글
fill_calc = PatternFill("solid", fgColor="F2F2F2")
fill_key = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")

wb = Workbook()

# ---------------------------------------------------------------- 코드표
ref = wb.active
ref.title = "코드표"
lists = {
    "A": ("병원", ["서울", "구리", "제주"]),
    "B": ("진료과", ["가정의학과", "내과", "외과", "소아청소년과", "산부인과", "정형외과", "마취통증의학과", "영상의학과", "응급의학과"]),
    "C": ("파견종류", ["수련 파견", "교육 파견", "지원 파견", "기타"]),
}
for col, (head, items) in lists.items():
    c = ref[f"{col}1"]
    c.value, c.font, c.fill, c.alignment, c.border = head, f_head, fill_head, center, border
    for i in range(2, 22):
        cell = ref[f"{col}{i}"]
        cell.font, cell.border = f_input, border
        if i - 2 < len(items):
            cell.value = items[i - 2]
    ref.column_dimensions[col].width = 20
ref["E1"].value = "※ 파란 글씨 칸이 드롭다운 목록입니다. 여기서 추가·수정하면 입력 목록과 요약이 같이 바뀝니다. (병원 목록은 기본자료의 '근무지' 표기와 맞추면 편합니다)"
ref["E1"].font = f_base

# ---------------------------------------------------------------- 기본자료 (업로드 시트)
bs = wb.create_sheet("기본자료", 0)
base_headers = ["현재근무 여부", "사번", "이름", "수련과목", "연차", "근무지", "수련개시일",
                "전공의등록번호", "의사면허번호", "성별", "생년월일", "비고"]
base_widths = [13, 11, 10, 14, 12, 10, 13, 15, 14, 7, 13, 18]
for i, (h, w) in enumerate(zip(base_headers, base_widths), 1):
    c = bs.cell(row=1, column=i, value=h)
    c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, border
    bs.column_dimensions[get_column_letter(i)].width = w
# 가상 예시 (업로드 양식과 동일 형식: 사번·등록번호=텍스트, 날짜=YYYYMMDD 숫자)
base_sample = [
    ("재직", "2261212", "김ㅇㅇ", "가정의학과", "레지던트1", "서울", 20260301, "11-1111", 111111, "여", 19999999, None),
    ("재직", "2261213", "이ㅇㅇ", "내과", "레지던트2", "서울", 20250301, "11-1112", 111112, "남", 19950702, None),
    ("재직", "2261214", "박ㅇㅇ", "소아청소년과", "레지던트1", "구리", 20260301, "11-1113", 111113, "여", 19970125, None),
    ("재직", "2261215", "최ㅇㅇ", "산부인과", "레지던트3", "서울", 20240301, "11-1114", 111114, "남", 19961109, None),
    ("재직", "2261216", "정ㅇㅇ", "마취통증의학과", "레지던트2", "서울", 20250301, "11-1115", 111115, "여", 19980530, None),
    ("재직", "222222", "김ㅇㅇ", "인턴", "인턴", "서울", 20260301, "22-2222", 222222, "여", 19990101, None),
    ("재직", "222223", "이ㅇㅇ", "인턴", "인턴", "서울", 20260301, "22-2223", 222223, "남", 19990202, None),
    ("재직", "222224", "박ㅇㅇ", "인턴", "인턴", "서울", 20260301, "22-2224", 222224, "남", 19990303, None),
]
for r in range(2, BLAST + 1):
    for c in range(1, 13):
        cell = bs.cell(row=r, column=c)
        cell.font, cell.border = f_input, border
        if c in (2, 8):
            cell.number_format = "@"
        if r - 2 < len(base_sample):
            v = base_sample[r - 2][c - 1]
            if v is not None:
                cell.value = v
bs.freeze_panes = "A2"
bs.auto_filter.ref = f"A1:L{BLAST}"

# ---------------------------------------------------------------- 전공의명단
ws = wb.create_sheet("전공의명단", 1)
# (머리글, 너비, 종류) 종류: calc=수식, lk=기본자료 조회, in=입력
cols = [
    ("번호", 6, "calc"), ("사번", 11, "key"), ("이름", 10, "lk"), ("성별", 6, "lk"),
    ("생년월일", 12, "lk"), ("수련과목", 14, "lk"), ("연차", 11, "lk"), ("근무지", 10, "lk"),
    ("수련개시일", 12, "lk"), ("전공의등록번호", 14, "lk"), ("의사면허번호", 13, "lk"),
    ("재직여부", 9, "lk"),
    ("현재 근무병원", 14, "in"), ("현재 진료과", 15, "in"), ("급여지급병원", 14, "in"),
    ("파견종류", 12, "in"), ("파견시작일", 12, "in"), ("파견종료일", 12, "in"),
    ("파견상태", 10, "calc"), ("파견기간(일)", 11, "calc"), ("종료까지(일)", 11, "calc"),
    ("비고", 22, "in"), ("기본자료행", 10, "calc"),
]
for i, (h, w, k) in enumerate(cols, 1):
    c = ws.cell(row=1, column=i, value=h)
    c.font, c.alignment, c.border = f_head, center, border
    c.fill = fill_head2 if k == "lk" else fill_head
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 26

BR = lambda col: f"기본자료!${col}$2:${col}${BLAST}"
# 기본자료 열 -> 명단 열
lookup = {3: "C", 4: "J", 5: "K", 6: "D", 7: "E", 8: "F", 9: "G", 10: "H", 11: "I", 12: "A"}
date_cols = {5, 9}


def pick(col, r):
    return f"INDEX({BR(col)},$W{r})"


def lk_formula(c, r):
    col = lookup[c]
    raw = pick(col, r)
    if c in date_cols:
        # YYYYMMDD 숫자 -> 날짜. 형식이 맞지 않으면(예: 19999999) 원문을 텍스트로 표시
        return (f'=IF($W{r}="","",IF({raw}="","",IF(NOT(ISNUMBER({raw})),{raw},'
                f'IF({raw}<1000000,{raw},'
                f'IF(AND({raw}>=10000101,INT(MOD({raw},10000)/100)>=1,INT(MOD({raw},10000)/100)<=12,'
                f'MOD({raw},100)>=1,MOD({raw},100)<=31),'
                f'DATE(INT({raw}/10000),INT(MOD({raw},10000)/100),MOD({raw},100)),{raw}&"")))))')
    return f'=IF($W{r}="","",IF({raw}="","",{raw}))'


# 가상 예시: 사번 + 파견 정보 (현재 근무병원, 진료과, 급여병원, 파견종류, 시작, 종료, 비고)
sample = [
    ("2261212", "서울", "가정의학과", "서울", "수련 파견", date(2026, 9, 1), date(2026, 10, 31), "예시 데이터"),
    ("2261213", "구리", "내과", "서울", "지원 파견", date(2026, 6, 1), date(2026, 12, 31), "예시 데이터"),
    ("2261214", "구리", "소아청소년과", "구리", None, None, None, "예시 데이터"),
    ("2261215", "제주", "산부인과", "서울", "교육 파견", date(2026, 1, 5), date(2026, 3, 31), "예시 데이터(종료됨)"),
    ("2261216", "서울", "마취통증의학과", "서울", "교육 파견", date(2026, 11, 1), date(2027, 1, 31), "예시 데이터(예정)"),
]
in_map = {13: 1, 14: 2, 15: 3, 16: 4, 17: 5, 18: 6, 22: 7}

for r in range(2, LAST + 1):
    ws.cell(row=r, column=1, value=f'=IF(B{r}="","",ROW()-1)')
    if r - 2 < len(sample):
        ws.cell(row=r, column=2, value=sample[r - 2][0])
        for c, idx in in_map.items():
            v = sample[r - 2][idx]
            if v is not None:
                ws.cell(row=r, column=c, value=v)
    for c in lookup:
        ws.cell(row=r, column=c, value=lk_formula(c, r))
    ws.cell(row=r, column=19, value=f'=IF(OR(B{r}="",Q{r}=""),"",IF(TODAY()<Q{r},"파견예정",IF(AND(R{r}<>"",TODAY()>R{r}),"파견종료","파견중")))')
    ws.cell(row=r, column=20, value=f'=IF(OR(Q{r}="",R{r}=""),"",R{r}-Q{r}+1)')
    ws.cell(row=r, column=21, value=f'=IF(AND(S{r}="파견중",R{r}<>""),R{r}-TODAY(),"")')
    # 사번이 숫자/텍스트 어느 쪽으로 입력·업로드돼도 찾도록 4가지 방식으로 매칭
    ws.cell(row=r, column=23, value=(
        f'=IF(B{r}="","",IFERROR(MATCH(B{r},{BR("B")},0),IFERROR(MATCH(B{r}&"",{BR("B")},0),'
        f'IFERROR(MATCH(VALUE(B{r}),{BR("B")},0),""))))'))
    for c in range(1, 24):
        cell = ws.cell(row=r, column=c)
        cell.border = border
        kind = cols[c - 1][2]
        cell.font = f_input if kind in ("key", "in") else f_base
        if kind in ("calc", "lk"):
            cell.fill = fill_calc
        if kind in ("calc", "lk", "key") and c != 22:
            cell.alignment = center
    for c in (5, 9, 17, 18):
        ws.cell(row=r, column=c).number_format = "yyyy-mm-dd"
        if c in (17, 18):
            ws.cell(row=r, column=c).alignment = center
    ws.cell(row=r, column=2).number_format = "@"
    ws.cell(row=r, column=20).number_format = "0"
    ws.cell(row=r, column=21).number_format = "0"

tab = Table(displayName="전공의", ref=f"A1:W{LAST}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
ws.add_table(tab)
ws.freeze_panes = "D2"
ws.column_dimensions["W"].hidden = False


def dv_list(rng_formula, target):
    dv = DataValidation(type="list", formula1=rng_formula, allow_blank=True, showErrorMessage=True,
                        errorTitle="목록 선택", error="코드표 시트의 목록에서 선택하세요.")
    ws.add_data_validation(dv)
    dv.add(target)


dv_list("=코드표!$A$2:$A$21", f"M2:M{LAST}")
dv_list("=코드표!$B$2:$B$21", f"N2:N{LAST}")
dv_list("=코드표!$A$2:$A$21", f"O2:O{LAST}")
dv_list("=코드표!$C$2:$C$21", f"P2:P{LAST}")
dv_date = DataValidation(type="date", operator="greaterThan", formula1="1900-01-01", allow_blank=True,
                         showErrorMessage=True, errorTitle="날짜 형식", error="yyyy-mm-dd 형식의 날짜를 입력하세요.")
ws.add_data_validation(dv_date)
dv_date.add(f"Q2:R{LAST}")
dv_end = DataValidation(type="custom", formula1='OR(R2="",Q2="",R2>=Q2)', allow_blank=True,
                        showErrorMessage=True, errorTitle="종료일 오류", error="종료일은 시작일 이후여야 합니다.")
ws.add_data_validation(dv_end)
dv_end.add(f"R2:R{LAST}")

# 조건부 서식 (먼저 추가한 규칙이 우선)
ws.conditional_formatting.add(f"B2:B{LAST}", FormulaRule(formula=['AND($B2<>"",$W2="")'], fill=PatternFill("solid", bgColor="FF0000"), font=Font(name=FONT, color="FFFFFF", bold=True)))
ws.conditional_formatting.add(f"B2:B{LAST}", FormulaRule(formula=['AND($B2<>"",COUNTIF($B$2:$B$%d,$B2)>1)' % LAST], fill=PatternFill("solid", bgColor="FFC000")))
ws.conditional_formatting.add(f"E2:E{LAST}", FormulaRule(formula=['ISTEXT($E2)'], font=Font(name=FONT, color="FF0000", bold=True)))
ws.conditional_formatting.add(f"I2:I{LAST}", FormulaRule(formula=['ISTEXT($I2)'], font=Font(name=FONT, color="FF0000", bold=True)))
ws.conditional_formatting.add(f"S2:U{LAST}", FormulaRule(formula=['$S2="파견종료"'], fill=PatternFill("solid", bgColor="D9D9D9"), font=Font(name=FONT, color="7F7F7F")))
ws.conditional_formatting.add(f"A2:V{LAST}", FormulaRule(formula=['AND($S2="파견중",$U2<>"",$U2<=요약!$E$2)'], fill=PatternFill("solid", bgColor="F8CBAD")))
ws.conditional_formatting.add(f"S2:S{LAST}", FormulaRule(formula=['$S2="파견중"'], fill=PatternFill("solid", bgColor="C6E0B4")))
ws.conditional_formatting.add(f"S2:S{LAST}", FormulaRule(formula=['$S2="파견예정"'], fill=PatternFill("solid", bgColor="BDD7EE")))
ws.conditional_formatting.add(f"O2:O{LAST}", FormulaRule(formula=['AND($O2<>"",$M2<>"",$O2<>$M2)'], fill=PatternFill("solid", bgColor="FFFF00")))

# ---------------------------------------------------------------- 요약
sm = wb.create_sheet("요약", 2)
sm["A1"].value, sm["A1"].font = "파견 현황 요약", f_title
sm["A2"].value, sm["A2"].font = "기준일", f_bold
sm["B2"].value, sm["B2"].number_format, sm["B2"].font = "=TODAY()", "yyyy-mm-dd", f_base
sm["D2"].value, sm["D2"].font = "종료 임박 기준(일)", f_bold
sm["E2"].value, sm["E2"].font, sm["E2"].fill = 30, f_input, fill_key
sm["F2"].value, sm["F2"].font = "← 변경 가능 (전공의명단 주황색 강조에도 반영)", f_base


def head(cell, text):
    sm[cell].value, sm[cell].font, sm[cell].fill, sm[cell].alignment, sm[cell].border = text, f_head, fill_head, center, border


def body(cell, value, bold=False):
    sm[cell].value, sm[cell].font, sm[cell].border = value, f_bold if bold else f_base, border
    sm[cell].alignment = center


N = "전공의명단"
rngS = f"{N}!$S$2:$S${LAST}"
head("A4", "파견상태"); head("B4", "인원")
for i, s in enumerate(["파견중", "파견예정", "파견종료"]):
    body(f"A{5+i}", s)
    body(f"B{5+i}", f"=COUNTIF({rngS},A{5+i})")
body("A8", "명단 등록 인원", True)
body("B8", f'=COUNTA({N}!$B$2:$B${LAST})', True)
body("A9", "종료 임박(파견중)")
body("B9", f'=COUNTIFS({rngS},"파견중",{N}!$U$2:$U${LAST},"<="&$E$2)')

head("D4", "병원"); head("E4", "현재 근무"); head("F4", "급여 지급"); head("G4", "파견중(근무기준)")
for i in range(20):
    r = 5 + i
    body(f"D{r}", f'=IF(코드표!A{2+i}="","",코드표!A{2+i})')
    body(f"E{r}", f'=IF(D{r}="","",COUNTIF({N}!$M$2:$M${LAST},D{r}))')
    body(f"F{r}", f'=IF(D{r}="","",COUNTIF({N}!$O$2:$O${LAST},D{r}))')
    body(f"G{r}", f'=IF(D{r}="","",COUNTIFS({N}!$M$2:$M${LAST},D{r},{rngS},"파견중"))')

head("I4", "진료과"); head("J4", "인원")
for i in range(20):
    r = 5 + i
    body(f"I{r}", f'=IF(코드표!B{2+i}="","",코드표!B{2+i})')
    body(f"J{r}", f'=IF(I{r}="","",COUNTIF({N}!$N$2:$N${LAST},I{r}))')

head("A12", "파견종류"); head("B12", "인원")
for i in range(8):
    r = 13 + i
    body(f"A{r}", f'=IF(코드표!C{2+i}="","",코드표!C{2+i})')
    body(f"B{r}", f'=IF(A{r}="","",COUNTIF({N}!$P$2:$P${LAST},A{r}))')

# 데이터 점검
head("A22", "데이터 점검"); head("B22", "건수")
body("A23", "기본자료 인원")
body("B23", f"=COUNTA({BR('B')})")
body("A24", "기본자료 재직 인원")
body("B24", f'=COUNTIF({BR("A")},"재직")')
body("A25", "기본자료에 없는 사번")
body("B25", f'=SUMPRODUCT(({N}!$B$2:$B${LAST}<>"")*({N}!$W$2:$W${LAST}=""))')
body("A26", "명단 중복 사번")
body("B26", f'=SUMPRODUCT(({N}!$B$2:$B${LAST}<>"")*(COUNTIF({N}!$B$2:$B${LAST},{N}!$B$2:$B${LAST})>1))')
body("A27", "날짜 형식 오류(생년월일/수련개시일)")
body("B27", f'=SUMPRODUCT(({N}!$E$2:$E${LAST}<>"")*1)-COUNT({N}!$E$2:$E${LAST})+SUMPRODUCT(({N}!$I$2:$I${LAST}<>"")*1)-COUNT({N}!$I$2:$I${LAST})')
sm.conditional_formatting.add("B25:B27", FormulaRule(formula=["B25>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
for col, w in zip("ABCDEFGHIJ", [30, 10, 3, 14, 12, 12, 16, 3, 18, 10]):
    sm.column_dimensions[col].width = w


# ---------------------------------------------------------------- 인턴근무계획 (업로드 시트)
PL = "인턴근무계획"
ip = wb.create_sheet(PL, 2)
IP_LAST = 202  # 데이터 3~202행 (200명)
ip.merge_cells("B1:K1")
ip["B1"].value = "2026년도 하반기 인턴근무계획표(9~2월)"
ip["B1"].font, ip["B1"].alignment = f_title, center
plan_headers = ["순번 ", "이름 ", "사번", "9월", "10월", "11월", "12월", "1월", "2월 ", "비고"]
plan_widths = [7, 10, 11, 14, 14, 14, 16, 14, 14, 10]
ip.column_dimensions["A"].width = 2
for i, (h, w) in enumerate(zip(plan_headers, plan_widths), 2):
    c = ip.cell(row=2, column=i, value=h)
    c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, border
    ip.column_dimensions[get_column_letter(i)].width = w
plan_sample = [
    (1, "김ㅇㅇ", 222222, "구리 안과", "구리 내과", "서울 핵방", "서울 재활의학과", "제주 산부", "서울 소아과", 2026.03),
    (2, "이ㅇㅇ", 222223, "서울 내과", "서울 외과", "구리 소아과", "구리 산부인과", "제주 안과", "서울 가정의학과", None),
    (3, "박ㅇㅇ", 222224, "제주 외과", "서울 내과", "서울 응급의학과", "구리 내과", "서울 소아과", "제주 산부", None),
]
for r in range(3, IP_LAST + 1):
    for c in range(2, 12):
        cell = ip.cell(row=r, column=c)
        cell.font, cell.border = f_input, border
        if r - 3 < len(plan_sample):
            v = plan_sample[r - 3][c - 2]
            if v is not None:
                cell.value = v
ip.freeze_panes = "D3"

# ---------------------------------------------------------------- 인턴현황 (조회·정렬용)
iv = wb.create_sheet("인턴현황", 3)
PR = lambda col: f"{PL}!${col}$3:${col}${IP_LAST}"
iv["A1"].value, iv["A1"].font = "인턴 근무 현황", f_title
iv["A2"].value, iv["A2"].font = "기준일", f_bold
iv["B2"].value, iv["B2"].number_format, iv["B2"].font = "=TODAY()", "yyyy-mm-dd", f_base
iv["A3"].value, iv["A3"].font = "계획 월 번호(자동)", Font(name=FONT, size=9, color="7F7F7F")
iv["E3"].value, iv["E3"].font = "이번/다음 달 위치→", Font(name=FONT, size=9, color="7F7F7F")
iv["F3"].value = '=IFERROR(MATCH(MONTH(TODAY()),$J$3:$O$3,0),"")'
iv["H3"].value = '=IFERROR(MATCH(MONTH(EDATE(TODAY(),1)),$J$3:$O$3,0),"")'
for c in ("F3", "H3"):
    iv[c].font, iv[c].alignment = Font(name=FONT, size=9, color="7F7F7F"), center
month_cols = "EFGHIJ"  # 인턴근무계획의 9월~2월 열
ih = {1: ("계획행", 7), 2: ("순번", 6), 3: ("이름", 10), 4: ("사번", 10), 5: ("기본자료 확인", 12),
      6: ('="이번 달("&IF($F$3="","계획 외",MONTH(TODAY())&"월")&") 병원"', 16), 7: ("이번 달 진료과", 16),
      8: ('="다음 달("&IF($H$3="","계획 외",MONTH(EDATE(TODAY(),1))&"월")&") 병원"', 16), 9: ("다음 달 진료과", 16),
      16: ("비고", 10)}
for k in range(6):
    ih[10 + k] = (f"=TRIM({PL}!{month_cols[k]}$2&\"\")", 9)
for c, (h, w) in ih.items():
    cell = iv.cell(row=5, column=c, value=h)
    cell.font, cell.fill, cell.alignment, cell.border = f_head, fill_head, center, border
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    iv.column_dimensions[get_column_letter(c)].width = w
iv.row_dimensions[5].height = 30
for k in range(6):
    col = get_column_letter(10 + k)
    iv[f"{col}3"].value = f'=IFERROR(VALUE(SUBSTITUTE(TRIM({col}5),"월","")),"")'
    iv[f"{col}3"].font, iv[f"{col}3"].alignment = Font(name=FONT, size=9, color="7F7F7F"), center


def cellraw(r, idx):  # 인턴근무계획 월 칸 원문 (공백 제거)
    return f'TRIM(INDEX({PL}!$E$3:$J${IP_LAST},$A{r},{idx})&"")'


def hosp(x):
    return f'IF({x}="","",IFERROR(LEFT({x},FIND(" ",{x})-1),{x}))'


def dept(x):
    return f'IF({x}="","",IFERROR(MID({x},FIND(" ",{x})+1,60),""))'


IV_FIRST = 6
IV_LAST = IV_FIRST + 200 - 1
BRng = lambda col: f"기본자료!${col}$2:${col}${BLAST}"
for i in range(200):
    r = IV_FIRST + i
    iv.cell(row=r, column=1, value=i + 1)  # 계획행 번호(상수): 정렬해도 같은 행끼리 유지
    iv.cell(row=r, column=3, value=f'=IFERROR(TRIM(INDEX({PR("C")},$A{r})&""),"")')
    iv.cell(row=r, column=2, value=f'=IF($C{r}="","",INDEX({PR("B")},$A{r}))')
    iv.cell(row=r, column=4, value=f'=IF($C{r}="","",INDEX({PR("D")},$A{r}))')
    iv.cell(row=r, column=5, value=(
        f'=IF($C{r}="","",IF(IFERROR(MATCH($D{r},{BRng("B")},0),IFERROR(MATCH($D{r}&"",{BRng("B")},0),'
        f'IFERROR(MATCH(VALUE($D{r}),{BRng("B")},0),0)))>0,"확인","미등록"))'))
    cur, nxt = cellraw(r, "$F$3"), cellraw(r, "$H$3")
    iv.cell(row=r, column=6, value=f'=IF(OR($C{r}="",$F$3=""),"",{hosp(cur)})')
    iv.cell(row=r, column=7, value=f'=IF(OR($C{r}="",$F$3=""),"",{dept(cur)})')
    iv.cell(row=r, column=8, value=f'=IF(OR($C{r}="",$H$3=""),"",{hosp(nxt)})')
    iv.cell(row=r, column=9, value=f'=IF(OR($C{r}="",$H$3=""),"",{dept(nxt)})')
    for k in range(6):
        raw = cellraw(r, k + 1)
        iv.cell(row=r, column=10 + k, value=f'=IF($C{r}="","",{hosp(raw)})')
    iv.cell(row=r, column=16, value=f'=IF($C{r}="","",IF(INDEX({PR("K")},$A{r})="","",INDEX({PR("K")},$A{r})))')
    for c in range(1, 17):
        cell = iv.cell(row=r, column=c)
        cell.border, cell.alignment = border, center
        cell.font = Font(name=FONT, size=9, color="7F7F7F") if c == 1 else f_base
iv.freeze_panes = "D6"
iv.auto_filter.ref = f"A5:P{IV_LAST}"

# 이번 달 열 강조 / 미등록 표시
iv.conditional_formatting.add(f"E{IV_FIRST}:E{IV_LAST}", FormulaRule(formula=[f'$E{IV_FIRST}="미등록"'], fill=PatternFill("solid", bgColor="FFC7CE")))
iv.conditional_formatting.add(f"J5:O{IV_LAST}", FormulaRule(formula=['J$3=MONTH(TODAY())'], fill=PatternFill("solid", bgColor="FFF2CC")))
iv.conditional_formatting.add(f"F{IV_FIRST}:G{IV_LAST}", FormulaRule(formula=['$F$3<>""'], fill=PatternFill("solid", bgColor="FFF2CC")))

# 월별 병원 인원 매트릭스
iv["R4"].value, iv["R4"].font = "월별 병원별 인턴 인원", f_bold
iv.cell(row=5, column=18, value="병원")
for k in range(6):
    iv.cell(row=5, column=19 + k, value=f"={get_column_letter(10 + k)}5")
for c in range(18, 25):
    cell = iv.cell(row=5, column=c)
    cell.font, cell.fill, cell.alignment, cell.border = f_head, fill_head, center, border
    iv.column_dimensions[get_column_letter(c)].width = 10 if c > 18 else 14
iv.column_dimensions["Q"].width = 3
for i in range(20):
    r = IV_FIRST + i
    iv.cell(row=r, column=18, value=f'=IF(코드표!A{2+i}="","",코드표!A{2+i})')
    for k in range(6):
        col = get_column_letter(10 + k)
        iv.cell(row=r, column=19 + k, value=f'=IF($R{r}="","",COUNTIF({col}${IV_FIRST}:{col}${IV_LAST},$R{r}))')
    for c in range(18, 25):
        cell = iv.cell(row=r, column=c)
        cell.font, cell.border, cell.alignment = f_base, border, center
r = IV_FIRST + 20
iv.cell(row=r, column=18, value="합계")
for k in range(6):
    col = get_column_letter(19 + k)
    iv.cell(row=r, column=19 + k, value=f"=SUM({col}{IV_FIRST}:{col}{r-1})")
for c in range(18, 25):
    cell = iv.cell(row=r, column=c)
    cell.font, cell.border, cell.alignment = f_bold, border, center
iv.cell(row=r + 1, column=18, value="※ 코드표에 없는 병원명은 합계에서 제외됩니다. 합계가 인원과 다르면 코드표를 확인하세요.").font = Font(name=FONT, size=9, color="7F7F7F")

# ---------------------------------------------------------------- 사용 안내
gd = wb.create_sheet("사용안내", 0)
lines = [
    ("전공의 파견 모니터링 사용 안내", f_title),
    ("", f_base),
    ("[전체 흐름]", f_bold),
    ("1) 기본자료 시트에 전공의 기본자료를 업로드(붙여넣기)합니다.  2) 전공의명단 시트에 사번을 입력하면 인적사항이 자동으로 채워집니다.  3) 현재 근무병원·진료과·급여지급병원·파견 정보를 입력합니다.", f_base),
    ("", f_base),
    ("[기본자료 업로드]", f_bold),
    ("• 기본자료 시트 A1:L1의 머리글 순서와 같은 양식(현재근무 여부·사번·이름·수련과목·연차·근무지·수련개시일·전공의등록번호·의사면허번호·성별·생년월일·비고)으로 A2부터 붙여넣으세요. 최대 500명.", f_base),
    ("• 열 순서는 반드시 유지해야 합니다(수식이 열 위치로 조회). 기본자료는 사번이 같으면 정렬 순서가 바뀌거나 갱신돼도 명단에 올바르게 연결됩니다.", f_base),
    ("• 날짜(수련개시일·생년월일)는 20260301 형식 숫자도 자동으로 날짜로 변환됩니다. 변환이 안 되는 값(예: 19999999)은 빨간 글씨 원문으로 표시되고 요약의 '날짜 형식 오류'에 집계됩니다.", f_base),
    ("• 기본자료 2~9행은 가상 예시(김ㅇㅇ 등)입니다. 삭제하고 실제 자료를 붙여넣으세요.", f_base),
    ("", f_base),
    ("[전공의명단 입력]", f_bold),
    ("• 파란 글씨 칸만 입력: 사번, 현재 근무병원, 현재 진료과, 급여지급병원, 파견종류, 파견시작일, 파견종료일, 비고.", f_base),
    ("• 초록 머리글(회색 칸)은 기본자료에서 자동 조회되며, 회색 칸(번호·파견상태·파견기간·종료까지·기본자료행)도 수식이므로 지우지 마세요.", f_base),
    ("• 모니터링할 전공의의 사번만 입력하면 됩니다(최대 200명). 기본자료에 없는 사번은 빨간색, 중복 사번은 주황색으로 표시됩니다.", f_base),
    ("• 파견이 없는 전공의는 파견종류·시작일·종료일을 비워 두세요. 종료일이 비어 있으면 종료 미정(파견중)으로 봅니다.", f_base),
    ("• 병원/진료과/파견종류 항목은 코드표 시트에서 관리합니다.", f_base),
    ("", f_base),
    ("[인턴 근무계획 (수련교육부)]", f_bold),
    ("• 인턴근무계획 시트에 인턴근무계획표를 업로드 양식 그대로(B1 제목, B2:K2 머리글, 3행부터 자료) 붙여넣으세요. 최대 200명. 3~5행은 가상 예시입니다.", f_base),
    ("• 월 칸은 '병원 진료과'(공백으로 구분, 예: 구리 내과) 형식이어야 병원/진료과로 자동 분리됩니다. 월 머리글이 바뀌면(예: 3~8월) 이번 달 위치가 자동으로 따라갑니다.", f_base),
    ("• 인턴현황 시트: 오늘 기준 이번 달·다음 달 근무 병원/진료과, 월별 병원, 기본자료 사번 확인(미등록 시 빨간색), 월별 병원별 인원표를 보여줍니다. 정렬·필터는 인턴현황에서 하세요(입력은 인턴근무계획에서만).", f_base),
    ("• 인턴을 기본자료에도 올려 두면(수련과목=인턴) 사번으로 인적사항을 확인할 수 있습니다.", f_base),
    ("", f_base),
    ("[정렬·필터]", f_bold),
    ("• 전공의명단 머리글의 ▼ 버튼으로 정렬/필터합니다 (예: 파견상태=파견중, 종료까지(일) 오름차순 → 종료 임박순). 사번·이름은 고정되어 가로 스크롤 시에도 보입니다.", f_base),
    ("", f_base),
    ("[자동 표시]", f_bold),
    ("• 파견상태: 오늘 기준 파견예정 / 파견중 / 파견종료.  주황색 행: 파견중이면서 종료까지 기준일수(요약 시트, 기본 30일) 이내.  회색: 파견종료.", f_base),
    ("• 급여지급병원 칸이 노란색: 현재 근무병원과 다름.", f_base),
    ("", f_base),
    ("[가정 사항]", f_bold),
    ("• 기본자료의 '근무지'(예: 서울)와 코드표의 병원명을 같은 표기로 맞추면 집계가 편합니다. 현재 근무병원은 파견 중 실제 근무하는 곳을 별도로 입력하는 칸입니다.", f_base),
    ("• 개인정보(생년월일·면허번호 등)가 포함되므로 파일 암호 설정과 접근 권한 관리를 권장합니다.", f_base),
]
for i, (t, f) in enumerate(lines, 1):
    gd.cell(row=i, column=1, value=t).font = f
gd.column_dimensions["A"].width = 160
gd.sheet_view.showGridLines = False

wb.save(OUT)
print("saved", OUT)
