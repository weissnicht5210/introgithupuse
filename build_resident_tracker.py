from datetime import date

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = "전공의_파견_모니터링.xlsx"
FONT = "맑은 고딕"
CAPACITY = 200  # 입력 가능한 전공의 행 수
LAST = CAPACITY + 1

f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_input = Font(name=FONT, size=10, color="0000FF")  # 입력값=파랑, 수식=검정
f_title = Font(name=FONT, size=14, bold=True)
fill_head = PatternFill("solid", fgColor="1F3864")
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
    "A": ("병원", ["본원", "A병원", "B병원", "C병원", "D병원"]),
    "B": ("진료과", ["내과", "외과", "소아청소년과", "산부인과", "정형외과", "마취통증의학과", "영상의학과", "응급의학과"]),
    "C": ("파견종류", ["수련 파견", "교육 파견", "지원 파견", "기타"]),
}
for col, (head, items) in lists.items():
    c = ref[f"{col}1"]
    c.value, c.font, c.fill, c.alignment, c.border = head, f_head, fill_head, center, border
    for i in range(2, 22):  # 20칸 확보 (목록 추가 가능)
        cell = ref[f"{col}{i}"]
        cell.font, cell.border = f_input, border
        if i - 2 < len(items):
            cell.value = items[i - 2]
    ref.column_dimensions[col].width = 20
ref["E1"].value = "※ 파란 글씨 칸이 드롭다운 목록입니다. 병원/진료과/파견종류를 이 표에서 추가·수정하면 입력 목록과 요약이 같이 바뀝니다."
ref["E1"].font = f_base

# ---------------------------------------------------------------- 전공의명단
ws = wb.create_sheet("전공의명단", 0)
headers = [
    ("번호", 6), ("이름", 11), ("사번", 12), ("생년월일", 12), ("입사장소", 14),
    ("의사면허번호", 14), ("전공의번호", 13), ("현재 근무병원", 15), ("현재 진료과", 16),
    ("급여지급병원", 15), ("파견종류", 13), ("파견시작일", 12), ("파견종료일", 12),
    ("파견상태", 11), ("파견기간(일)", 11), ("종료까지(일)", 11), ("비고", 24),
]
for i, (h, w) in enumerate(headers, 1):
    c = ws.cell(row=1, column=i, value=h)
    c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, border
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 24

# 예시 데이터(가상). 삭제하고 실제 데이터를 입력하세요.
sample = [
    ("김가상", "A10001", date(1996, 3, 14), "본원", "M-100001", "R-2024-001", "본원", "내과", "본원", "수련 파견", date(2026, 9, 1), date(2026, 10, 31), "예시 데이터"),
    ("이예시", "A10002", date(1995, 7, 2), "본원", "M-100002", "R-2024-002", "A병원", "외과", "본원", "지원 파견", date(2026, 6, 1), date(2026, 12, 31), "예시 데이터"),
    ("박샘플", "A10003", date(1997, 1, 25), "A병원", "M-100003", "R-2025-003", "A병원", "소아청소년과", "A병원", "", None, None, "예시 데이터"),
    ("최테스트", "A10004", date(1996, 11, 9), "본원", "M-100004", "R-2025-004", "B병원", "산부인과", "본원", "교육 파견", date(2026, 1, 5), date(2026, 3, 31), "예시 데이터(종료됨)"),
    ("정미래", "A10005", date(1998, 5, 30), "본원", "M-100005", "R-2025-005", "본원", "마취통증의학과", "본원", "교육 파견", date(2026, 11, 1), date(2027, 1, 31), "예시 데이터(예정)"),
]
# 입력 열 -> 시트 열 매핑
cols_in = {1: 2, 2: 3, 3: 4, 4: 5, 5: 6, 6: 7, 7: 8, 8: 9, 9: 10, 10: 11, 11: 12, 12: 13, 13: 17}

for r in range(2, LAST + 1):
    ws.cell(row=r, column=1, value=f'=IF(B{r}="","",ROW()-1)')
    if r - 2 < len(sample):
        for k, col in cols_in.items():
            v = sample[r - 2][k - 1]
            if v not in ("", None):
                ws.cell(row=r, column=col, value=v)
    ws.cell(row=r, column=14, value=f'=IF(OR(B{r}="",L{r}=""),"",IF(TODAY()<L{r},"파견예정",IF(AND(M{r}<>"",TODAY()>M{r}),"파견종료","파견중")))')
    ws.cell(row=r, column=15, value=f'=IF(OR(L{r}="",M{r}=""),"",M{r}-L{r}+1)')
    ws.cell(row=r, column=16, value=f'=IF(AND(N{r}="파견중",M{r}<>""),M{r}-TODAY(),"")')
    for c in range(1, 18):
        cell = ws.cell(row=r, column=c)
        cell.border = border
        calc = c in (1, 14, 15, 16)
        cell.font = f_base if calc else f_input
        if calc:
            cell.fill = fill_calc
            cell.alignment = center
        elif c in (4, 12, 13):
            cell.number_format = "yyyy-mm-dd"
            cell.alignment = center
    ws.cell(row=r, column=15).number_format = "0"
    ws.cell(row=r, column=16).number_format = "0"
    for c in (3, 6, 7):  # 사번/면허/전공의번호는 텍스트로 취급
        ws.cell(row=r, column=c).number_format = "@"

tab = Table(displayName="전공의", ref=f"A1:Q{LAST}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
ws.add_table(tab)
ws.freeze_panes = "C2"

# 드롭다운 / 날짜 검증
def dv_list(rng_formula, target):
    dv = DataValidation(type="list", formula1=rng_formula, allow_blank=True, showErrorMessage=True,
                        errorTitle="목록 선택", error="코드표 시트의 목록에서 선택하세요.")
    ws.add_data_validation(dv)
    dv.add(target)

dv_list("=코드표!$A$2:$A$21", f"E2:E{LAST}")
dv_list("=코드표!$A$2:$A$21", f"H2:H{LAST}")
dv_list("=코드표!$B$2:$B$21", f"I2:I{LAST}")
dv_list("=코드표!$A$2:$A$21", f"J2:J{LAST}")
dv_list("=코드표!$C$2:$C$21", f"K2:K{LAST}")
dv_date = DataValidation(type="date", operator="greaterThan", formula1="1900-01-01", allow_blank=True,
                         showErrorMessage=True, errorTitle="날짜 형식", error="yyyy-mm-dd 형식의 날짜를 입력하세요.")
ws.add_data_validation(dv_date)
dv_date.add(f"D2:D{LAST}")
dv_date.add(f"L2:M{LAST}")
dv_end = DataValidation(type="custom", formula1='OR(M2="",L2="",M2>=L2)', allow_blank=True,
                        showErrorMessage=True, errorTitle="종료일 오류", error="종료일은 시작일 이후여야 합니다.")
ws.add_data_validation(dv_end)
dv_end.add(f"M2:M{LAST}")

# 조건부 서식 (모니터링)
rng = f"A2:Q{LAST}"
ws.conditional_formatting.add(f"N2:P{LAST}", FormulaRule(formula=['$N2="파견종료"'], fill=PatternFill("solid", bgColor="D9D9D9"), font=Font(name=FONT, color="7F7F7F")))
ws.conditional_formatting.add(f"A2:Q{LAST}", FormulaRule(formula=['AND($N2="파견중",$P2<>"",$P2<=요약!$E$2)'], fill=PatternFill("solid", bgColor="F8CBAD")))
ws.conditional_formatting.add(f"N2:N{LAST}", FormulaRule(formula=['$N2="파견중"'], fill=PatternFill("solid", bgColor="C6E0B4")))
ws.conditional_formatting.add(f"N2:N{LAST}", FormulaRule(formula=['$N2="파견예정"'], fill=PatternFill("solid", bgColor="BDD7EE")))
# 급여지급병원 != 현재 근무병원 → 노란 표시
ws.conditional_formatting.add(f"J2:J{LAST}", FormulaRule(formula=['AND($J2<>"",$H2<>"",$J2<>$H2)'], fill=PatternFill("solid", bgColor="FFFF00")))

# ---------------------------------------------------------------- 요약
sm = wb.create_sheet("요약", 1)
sm["A1"].value, sm["A1"].font = "파견 현황 요약", f_title
sm["A2"].value, sm["A2"].font = "기준일", f_bold
sm["B2"].value, sm["B2"].number_format, sm["B2"].font = "=TODAY()", "yyyy-mm-dd", f_base
sm["D2"].value, sm["D2"].font = "종료 임박 기준(일)", f_bold
sm["E2"].value, sm["E2"].font, sm["E2"].fill = 30, f_input, fill_key
sm["F2"].value, sm["F2"].font = "← 변경 가능 (전공의명단 주황색 강조에도 반영)", f_base

def head(cell, text):
    sm[cell].value, sm[cell].font, sm[cell].fill, sm[cell].alignment, sm[cell].border = text, f_head, fill_head, center, border

def body(cell, value, fmt=None):
    sm[cell].value, sm[cell].font, sm[cell].border = value, f_base, border
    sm[cell].alignment = center
    if fmt:
        sm[cell].number_format = fmt

N = "전공의명단"
rngN = f"{N}!$N$2:$N${LAST}"
# 1) 파견상태별
head("A4", "파견상태"); head("B4", "인원")
for i, s in enumerate(["파견중", "파견예정", "파견종료"]):
    body(f"A{5+i}", s)
    body(f"B{5+i}", f'=COUNTIF({rngN},A{5+i})')
body("A8", "전체 인원"); sm["A8"].font = f_bold
body("B8", f'=COUNTA({N}!$B$2:$B${LAST})'); sm["B8"].font = f_bold
body("A9", "종료 임박(파견중)")
body("B9", f'=COUNTIFS({rngN},"파견중",{N}!$P$2:$P${LAST},"<="&$E$2)')

# 2) 병원별: 현재 근무 / 급여지급 / 파견중 인원
head("D4", "병원"); head("E4", "현재 근무"); head("F4", "급여 지급"); head("G4", "파견중(근무기준)")
for i in range(20):
    r = 5 + i
    body(f"D{r}", f'=IF(코드표!A{2+i}="","",코드표!A{2+i})')
    body(f"E{r}", f'=IF(D{r}="","",COUNTIF({N}!$H$2:$H${LAST},D{r}))')
    body(f"F{r}", f'=IF(D{r}="","",COUNTIF({N}!$J$2:$J${LAST},D{r}))')
    body(f"G{r}", f'=IF(D{r}="","",COUNTIFS({N}!$H$2:$H${LAST},D{r},{rngN},"파견중"))')

# 3) 진료과별
head("I4", "진료과"); head("J4", "인원")
for i in range(20):
    r = 5 + i
    body(f"I{r}", f'=IF(코드표!B{2+i}="","",코드표!B{2+i})')
    body(f"J{r}", f'=IF(I{r}="","",COUNTIF({N}!$I$2:$I${LAST},I{r}))')

# 4) 파견종류별
head("A12", "파견종류"); head("B12", "인원")
for i in range(8):
    r = 13 + i
    body(f"A{r}", f'=IF(코드표!C{2+i}="","",코드표!C{2+i})')
    body(f"B{r}", f'=IF(A{r}="","",COUNTIF({N}!$K$2:$K${LAST},A{r}))')
for col, w in zip("ABCDEFGHIJ", [20, 10, 3, 14, 12, 12, 16, 3, 18, 10]):
    sm.column_dimensions[col].width = w

# ---------------------------------------------------------------- 사용 안내
gd = wb.create_sheet("사용안내", 0)
lines = [
    ("전공의 파견 모니터링 사용 안내", f_title),
    ("", f_base),
    ("[입력]", f_bold),
    ("• 파란 글씨 칸만 입력하세요. 회색 칸(번호·파견상태·파견기간·종료까지)은 자동 계산 수식이므로 지우지 마세요.", f_base),
    ("• 전공의명단 시트의 2~6행은 가상 예시입니다. 삭제하고 실제 정보를 입력하세요. (수식 열은 200행까지 미리 채워져 있습니다)", f_base),
    ("• 입력 가능 인원: 200명 (더 필요하면 표 마지막 행 아래로 수식 열을 복사해 늘리세요).", f_base),
    ("• 입사장소/근무병원/급여지급병원/진료과/파견종류는 드롭다운입니다. 항목은 코드표 시트에서 관리합니다.", f_base),
    ("• 날짜는 yyyy-mm-dd 형식. 파견종료일은 시작일 이후여야 합니다. 파견이 없는 전공의는 파견종류·시작일·종료일을 비워 두세요.", f_base),
    ("", f_base),
    ("[정렬·필터]", f_bold),
    ("• 전공의명단 머리글의 ▼ 버튼으로 정렬/필터합니다 (예: 파견상태=파견중, 종료까지(일) 오름차순 → 종료 임박순).", f_base),
    ("• 이름·번호 열은 고정(틀 고정)되어 가로 스크롤 시에도 보입니다.", f_base),
    ("", f_base),
    ("[자동 표시]", f_bold),
    ("• 파견상태: 오늘 기준 파견예정 / 파견중 / 파견종료 (파일을 열 때마다 TODAY()로 갱신).", f_base),
    ("• 주황색 행: 파견중이면서 종료까지 요약 시트의 기준일수(기본 30일) 이내.  회색: 파견종료.  노란색 급여지급병원: 현재 근무병원과 다름.", f_base),
    ("", f_base),
    ("[가정 사항]", f_bold),
    ("• '파견시작정보/종료정보'는 시작일·종료일로 구성했습니다. 파견처 병원 등 항목이 더 필요하면 열을 추가하세요.", f_base),
    ("• 종료일을 비워 두면 종료 미정(진행 중)으로 보고 파견중으로 표시합니다.", f_base),
    ("• 개인정보(생년월일·면허번호 등)가 포함되므로 파일 암호 설정과 접근 권한 관리를 권장합니다.", f_base),
]
for i, (t, f) in enumerate(lines, 1):
    gd.cell(row=i, column=1, value=t).font = f
gd.column_dimensions["A"].width = 130
gd.sheet_view.showGridLines = False

wb.save(OUT)
print("saved", OUT)
