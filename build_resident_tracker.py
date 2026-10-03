import os
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
PL = "인턴근무계획"
PLAN_LAST = 202  # 인턴근무계획 데이터 3~202행 (200명)
PL2 = "파견수련계획"
PLAN2_LAST = 203  # 파견수련계획 데이터 4~203행 (200건)
INTERN_CAP = 200  # 인턴현황 행 수
# 구글 시트용 빌드 (GSHEETS=1): 업로드 크기를 줄이고, 다른 시트를 참조하는 조건부서식·검증 수식을 INDIRECT로 감싼다
GS = os.environ.get("GSHEETS") == "1"
if GS:
    OUT = "전공의_파견_모니터링_구글시트용.xlsx"
    CAPACITY, BASE_CAP, INTERN_CAP = 60, 300, 60
    LAST, BLAST = CAPACITY + 1, BASE_CAP + 1
    PLAN_LAST, PLAN2_LAST = 2 + INTERN_CAP, 3 + 80


def xref(a):
    return f'INDIRECT("{a}")' if GS else a

HOME_SPECIAL = "구리병원"  # 15일 규칙이 적용되는 소속(입사장소) 병원

f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_input = Font(name=FONT, size=10, color="0000FF")  # 입력값=파랑, 수식=검정
f_title = Font(name=FONT, size=14, bold=True)
f_note = Font(name=FONT, size=9, color="7F7F7F")
fill_head = PatternFill("solid", fgColor="1F3864")
fill_head2 = PatternFill("solid", fgColor="548235")  # 자동 조회·계산 열 머리글
fill_calc = PatternFill("solid", fgColor="F2F2F2")
fill_key = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")
wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

HOSPITALS = ["서울병원", "인천병원", "구리병원", "창원병원", "제주병원", "오산병원"]
DEPTS = ["내과", "소아청소년과", "신경과", "정신건강의학과", "피부과", "외과", "심장혈관흉부외과",
         "정형외과", "신경외과", "성형외과", "산부인과", "안과", "이비인후과", "비뇨의학과",
         "재활의학과", "마취통증의학과", "영상의학과", "진단검사의학과", "병리과", "가정의학과",
         "응급의학과", "핵의학과", "직업환경의학과", "수련교육부(인턴)"]
INTERN_ONLY = ["통합", "핵/방"]  # 인턴 순환 전용 구분 (통합=과 미지정, 핵/방=핵의학과+방사선종양학과)
DISPATCH = ["모자", "통합수련", "다기관"]  # 파견근거(양식: 모자/통합수련) + 다기관
ALIASES = {  # 약칭 -> 정식 (정식명칭 자체는 자동 등록)
    "소아과": "소아청소년과", "소청": "소아청소년과", "소청과": "소아청소년과",
    "정신과": "정신건강의학과", "정신": "정신건강의학과",
    "흉부외과": "심장혈관흉부외과", "흉외": "심장혈관흉부외과", "심장혈관흉부": "심장혈관흉부외과",
    "정형": "정형외과", "신외": "신경외과", "성형": "성형외과", "산부": "산부인과",
    "이비": "이비인후과", "이비인후": "이비인후과", "비뇨": "비뇨의학과", "비뇨기과": "비뇨의학과",
    "재활": "재활의학과", "마취": "마취통증의학과", "마통": "마취통증의학과", "영상": "영상의학과",
    "진검": "진단검사의학과", "진단검사": "진단검사의학과", "병리": "병리과",
    "가정": "가정의학과", "가정의학": "가정의학과", "응급": "응급의학과",
    "핵의학": "핵의학과", "핵": "핵의학과", "직업환경": "직업환경의학과", "직환": "직업환경의학과",
    "수련교육부": "수련교육부(인턴)", "인턴": "수련교육부(인턴)",
    "핵방": "핵/방", "핵/방": "핵/방", "통합": "통합",
}

HOSP_ALIAS = {  # 병원 표기 -> 정식 병원 (정식명·지역명은 자동 등록). 기관명은 실제 표기에 맞게 코드표에서 추가
    "한양대학교병원": "서울병원", "한양대병원": "서울병원", "한양대서울병원": "서울병원",
    "한양대구리병원": "구리병원", "에스중앙병원": "제주병원", "에스중앙": "제주병원",
}

wb = Workbook()

# ---------------------------------------------------------------- 코드표
ref = wb.active
ref.title = "코드표"
HDR = {"A": "병원", "B": "근무과(진료과)", "C": "파견근거(종류)", "D": "인턴 전용 구분", "F": "진료과 약칭", "G": "정식 명칭",
       "J": "병원 표기(약칭·기관명)", "K": "정식 병원"}
for col, text in HDR.items():
    c = ref[f"{col}1"]
    c.value, c.font, c.fill, c.alignment, c.border = text, f_head, fill_head, wrap_center, border
    ref.column_dimensions[col].width = 20
for col in "EHIL":
    ref.column_dimensions[col].width = 3
ref.row_dimensions[1].height = 30
alias_rows = [(d, d) for d in DEPTS] + list(ALIASES.items())
hosp_alias_rows = [(h, h) for h in HOSPITALS] + [(h[:-2], h) for h in HOSPITALS] + list(HOSP_ALIAS.items())
for col, items, n in (("A", HOSPITALS, 20), ("B", DEPTS, 40), ("C", DISPATCH, 10), ("D", INTERN_ONLY, 5)):
    for i in range(n):
        c = ref[f"{col}{2+i}"]
        c.font, c.border = f_input, border
        if i < len(items):
            c.value = items[i]
for i in range(100):
    for col in "FGJK":
        ref[f"{col}{2+i}"].font = f_input
        ref[f"{col}{2+i}"].border = border
    if i < len(alias_rows):
        ref[f"F{2+i}"].value, ref[f"G{2+i}"].value = alias_rows[i]
    if i < len(hosp_alias_rows):
        ref[f"J{2+i}"].value, ref[f"K{2+i}"].value = hosp_alias_rows[i]
notes = [
    "※ 파란 글씨 칸은 모두 수정·추가할 수 있습니다. 새 병원/진료과가 생기면 여기에 먼저 추가하세요.",
    "※ 병원 표기(J→K): 인턴근무계획의 지역 단어와 파견수련계획의 모병원명·파견병원은 이 표로 정식 병원(A열)으로 바뀝니다.",
    "   (예: 한양대학교병원→서울병원, 한양대구리병원→구리병원, '제주 에스중앙 내과'의 '제주'→제주병원). 인천·창원·오산 등 실제 기관명은 J·K열에 추가하세요.",
    "※ 진료과 약칭(F→G): 인턴근무계획 월 칸의 마지막 단어와 파견수련계획의 파견과목을 정식 명칭으로 바꿉니다. (예: 산부→산부인과)",
    "※ 파견근거(C열): 양식의 모자 / 통합수련, 그리고 다기관. 새 값이 있으면 추가하세요.",
]
for i, t in enumerate(notes):
    ref[f"M{1+i}"].value, ref[f"M{1+i}"].font = t, f_base

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
    ("재직", "2261214", "박ㅇㅇ", "내과", "레지던트3", "서울", 20240301, "21-1111", 111113, "여", 19970125, None),
    ("재직", "2261215", "최ㅇㅇ", "산부인과", "레지던트3", "서울", 20240301, "11-1114", 111114, "남", 19961109, None),
    ("재직", "2261216", "정ㅇㅇ", "마취통증의학과", "레지던트2", "구리", 20250301, "11-1115", 111115, "여", 19980530, None),
    ("재직", "222222", "김ㅇㅇ", "인턴", "인턴", "서울", 20260301, "22-2222", 222222, "여", 19990101, None),
    ("재직", "222223", "이ㅇㅇ", "인턴", "인턴", "서울", 20260301, "22-2223", 222223, "남", 19990202, None),
    ("재직", "222224", "박ㅇㅇ", "인턴", "인턴", "구리", 20260301, "22-2224", 222224, "남", 19990303, None),
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

# ---------------------------------------------------------------- 인턴근무계획 (업로드 시트)
ip = wb.create_sheet(PL, 1)
ip.merge_cells("B1:K1")
ip["B1"].value = "2026년도 하반기 인턴근무계획표(9~2월)"
ip["B1"].font, ip["B1"].alignment = f_title, center
plan_headers = ["순번 ", "이름 ", "사번", "9월", "10월", "11월", "12월", "1월", "2월 ", "비고"]
plan_widths = [7, 10, 11, 18, 18, 18, 18, 18, 18, 10]
ip.column_dimensions["A"].width = 2
ip.column_dimensions["L"].width = 3
for i, (h, w) in enumerate(zip(plan_headers, plan_widths), 2):
    c = ip.cell(row=2, column=i, value=h)
    c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, border
    ip.column_dimensions[get_column_letter(i)].width = w
plan_sample = [
    (1, "김ㅇㅇ", 222222, "구리 안과", "구리 내과", "서울 핵방", "서울 재활의학과", "제주 산부", "서울 소아과", 2026.03),
    (2, "이ㅇㅇ", 222223, "서울 내과", "서울 외과", "구리 통합", "인천 산부인과", "제주 에스중앙 내과", "서울 가정의학과", 2026.03),
    (3, "박ㅇㅇ", 222224, "제주 외과", "서울 내과", "서울 응급", "구리 내과", "서울 소아과", "제주 산부", 2026.03),
]
for r in range(3, PLAN_LAST + 1):
    for c in range(2, 12):
        cell = ip.cell(row=r, column=c)
        cell.font, cell.border = f_input, border
        if r - 3 < len(plan_sample):
            v = plan_sample[r - 3][c - 2]
            if v is not None:
                cell.value = v
ip.freeze_panes = "D3"

MONTH_SRC = "EFGHIJ"   # 업로드 월 칸 (9월~2월)
MONTH_NORM = "MNOPQR"  # 정규화 결과 열


def norm_formula(src, r):
    raw = f'TRIM({src}{r}&"")'
    t1 = f'LEFT({raw},FIND(" ",{raw})-1)'
    hn = f'INDEX(코드표!$K$2:$K$101,MATCH({t1},코드표!$J$2:$J$101,0))'
    dt = f'TRIM(RIGHT(SUBSTITUTE({raw}," ",REPT(" ",60)),60))'
    fl = "코드표!$F$2:$F$101"
    return (f'=IFERROR(IF({raw}="","",IF(AND(ISNUMBER(FIND(" ",{raw})),ISNUMBER(MATCH({t1},코드표!$J$2:$J$101,0)),'
            f'ISNUMBER(MATCH({dt},{fl},0))),{hn}&" "&INDEX(코드표!$G$2:$G$101,MATCH({dt},{fl},0)),"⚠"&{raw})),'
            f'"⚠"&TRIM({src}{r}&""))')


ip["M1"].value, ip["M1"].font = "↓ 자동 정규화 (수정 금지) — '병원 진료과'로 통일, 인식 불가는 ⚠ 표시", f_note
for k, col in enumerate(MONTH_NORM):
    c = ip[f"{col}2"]
    c.value = f'=TRIM({MONTH_SRC[k]}$2&"")&" 정규화"'
    c.font, c.fill, c.alignment, c.border = f_head, fill_head2, wrap_center, border
    ip.column_dimensions[col].width = 20
for col, h, w in (("S", "입사일(비고)", 13), ("T", "입력오류(건)", 11)):
    c = ip[f"{col}2"]
    c.value, c.font, c.fill, c.alignment, c.border = h, f_head, fill_head2, wrap_center, border
    ip.column_dimensions[col].width = w
ip.row_dimensions[2].height = 28
for r in range(3, PLAN_LAST + 1):
    for k, col in enumerate(MONTH_NORM):
        ip[f"{col}{r}"].value = norm_formula(MONTH_SRC[k], r)
    v = f'VALUE(K{r}&"")'
    mo = f'ROUND(({v}-INT({v}))*100,0)'
    ip[f"S{r}"].value = (f'=IF(K{r}="","",IFERROR(IF(AND({mo}>=1,{mo}<=12,INT({v})>=1900),'
                         f'DATE(INT({v}),{mo},1),K{r}&""),K{r}&""))')
    ip[f"T{r}"].value = f'=COUNTIF(M{r}:R{r},"⚠*")'
    for c in list(range(13, 19)) + [19, 20]:
        cell = ip.cell(row=r, column=c)
        cell.font, cell.fill, cell.border, cell.alignment = f_base, fill_calc, border, center
    ip[f"S{r}"].number_format = "yyyy-mm-dd"

# 입력 검증: 모르는 병원/진료과 약칭이면 오류 팝업 (직접 입력할 때)
t = 'TRIM(E3)'
t1 = f'LEFT({t},FIND(" ",{t})-1)'
dv_formula = (f'AND(ISNUMBER(FIND(" ",{t})),ISNUMBER(MATCH({t1},{xref("코드표!$J$2:$J$101")},0)),'
              f'ISNUMBER(MATCH(TRIM(RIGHT(SUBSTITUTE({t}," ",REPT(" ",60)),60)),{xref("코드표!$F$2:$F$101")},0)))')
dv_plan = DataValidation(type="custom", formula1=dv_formula, allow_blank=True, showErrorMessage=True,
                         errorStyle="stop", errorTitle="등록되지 않은 병원/진료과",
                         error="'지역 진료과' 형식(예: 서울 내과, 제주 산부)으로 입력하세요.\n"
                               "처음 보는 지역이나 진료과(약칭)라면 먼저 '코드표' 시트에 추가한 뒤 다시 입력하세요.")
ip.add_data_validation(dv_plan)
dv_plan.add(f"E3:J{PLAN_LAST}")
ip.conditional_formatting.add(f"E3:J{PLAN_LAST}", FormulaRule(formula=['LEFT(M3,1)="⚠"'], fill=PatternFill("solid", bgColor="FFC7CE")))
ip.conditional_formatting.add(f"M3:R{PLAN_LAST}", FormulaRule(formula=['LEFT(M3,1)="⚠"'], fill=PatternFill("solid", bgColor="FFC7CE"), font=Font(name=FONT, color="C00000", bold=True)))

# ---------------------------------------------------------------- 파견수련계획 (업로드 시트)
dp = wb.create_sheet(PL2, 2)
dp_headers = ["연번", "모병원명", "수련\n종별\n(I/R)", "수련\n과목", "연차", "성명", "전공의\n등록번호",
              "파견시작일", "파견종료일", "파견일수", "파견병원", "파견\n과목", "파견근거\n(모자/통합수련)", "비고"]
dp_widths = [6, 16, 8, 14, 9, 10, 11, 12, 12, 9, 16, 14, 14, 20]
for i, (h, w) in enumerate(zip(dp_headers, dp_widths), 1):
    dp.merge_cells(start_row=1, start_column=i, end_row=3, end_column=i)
    c = dp.cell(row=1, column=i, value=h)
    c.font, c.fill, c.alignment = f_head, fill_head, wrap_center
    for rr in (1, 2, 3):
        dp.cell(row=rr, column=i).border = border
        dp.cell(row=rr, column=i).fill = fill_head
    dp.column_dimensions[get_column_letter(i)].width = w
dp.column_dimensions["O"].width = 30
dp.column_dimensions["P"].width = 3
dp_sample = [
    ("한양대학교병원", "R", "내과", 3, "박ㅇㅇ", "21-1111", date(2026, 5, 2), date(2026, 8, 31), "한양대구리병원", "내과", "모자", None, "2026.09.01-2026.09.30 다기관 파견"),
    ("한양대학교병원", "R", "내과", 3, "박ㅇㅇ", "21-1111", date(2026, 10, 1), date(2026, 10, 31), "한양대구리병원", "내과", "모자", None, None),
    ("한양대학교병원", "R", "가정의학과", 1, "김ㅇㅇ", "11-1111", date(2026, 10, 21), date(2026, 12, 31), "한양대구리병원", "가정의학과", "통합수련", None, None),
    ("한양대학교병원", "R", "내과", 2, "이ㅇㅇ", "11-1112", date(2026, 10, 5), date(2026, 12, 31), "에스중앙병원", "내과", "통합수련", None, None),
    ("한양대구리병원", "R", "마취통증의학과", 2, "정ㅇㅇ", "11-1115", date(2026, 10, 18), date(2026, 12, 31), "한양대학교병원", "마취통증의학과", "모자", None, None),
    ("한양대학교병원", "R", "산부인과", 3, "최ㅇㅇ", "11-1114", date(2026, 1, 5), date(2026, 3, 31), "창원", "산부인과", "다기관", None, None),
    ("한양대구리병원", "I", "인턴", "인턴", "박ㅇㅇ", "22-2224", date(2026, 11, 1), date(2026, 11, 30), "한양대학교병원", "응급의학과", "모자", None, None),
]
# 업로드 열 -> 샘플 인덱스 (B~N), 파견일수(J)는 수식
for r in range(4, PLAN2_LAST + 1):
    k = r - 4
    dp.cell(row=r, column=1, value=f'=IF(F{r}="","",ROW()-3)')
    dp.cell(row=r, column=10, value=f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r})),I{r}-H{r}+1,"")')
    if k < len(dp_sample):
        rec = dp_sample[k]
        for c, idx in zip([2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15], range(13)):
            v = rec[idx]
            if v is not None:
                dp.cell(row=r, column=c, value=v)
    for c in range(1, 16):
        cell = dp.cell(row=r, column=c)
        cell.font, cell.border = f_input, border
        if c in (1, 3, 5, 10):
            cell.alignment = center
    for c in (8, 9):
        dp.cell(row=r, column=c).number_format = "yyyy-mm-dd"
        dp.cell(row=r, column=c).alignment = center
    dp.cell(row=r, column=7).number_format = "@"
dp.freeze_panes = "H4"

# 자동 정규화 열 (Q~AA): 입력 영역(A~O) 오른쪽, 수정 금지
BOM, EOM_ = "요약!$B$3", "EOMONTH(요약!$B$3,0)"
dp["Q1"].value, dp["Q1"].font = "↓ 자동 계산 (수정 금지) — 병원/과목은 코드표 기준으로 통일, 인식 불가는 ⚠ 표시", f_note
helpers = [("Q", "키(등록번호)", 12), ("R", "모병원(정규화)", 12), ("S", "파견병원(정규화)", 13), ("T", "파견근거(정규화)", 11),
           ("U", "파견과목(정규화)", 15), ("V", "파견일수(계산)", 9), ("W", "급여기준월 겹침일수", 11), ("X", "상태", 10),
           ("Y", "키|상태", 16), ("Z", "키|기준월", 14), ("AA", "입력 오류", 22),
           ("AB", "비고 다기관 시작", 12), ("AC", "비고 다기관 종료", 12), ("AD", "비고 다기관 기준월 겹침일수", 12), ("AE", "키|다기관", 14)]
for col, h, w in helpers:
    c = dp[f"{col}3"]
    c.value, c.font, c.fill, c.alignment, c.border = h, f_head, fill_head2, wrap_center, border
    dp.column_dimensions[col].width = w
    dp[f"{col}2"].fill = fill_head2
    dp[f"{col}1"].fill = PatternFill(fill_type=None)
dp["Q1"].fill = PatternFill(fill_type=None)
dp.row_dimensions[3].height = 30


def hn_cell(col, r):
    x = f'TRIM({col}{r}&"")'
    return f'IF({x}="","",IFERROR(INDEX(코드표!$K$2:$K$101,MATCH({x},코드표!$J$2:$J$101,0)),"⚠"&{x}))'


for r in range(4, PLAN2_LAST + 1):
    dp[f"Q{r}"].value = f'=IF(TRIM(G{r}&"")="","",TRIM(G{r}&""))'
    dp[f"R{r}"].value = "=" + hn_cell("B", r)
    dp[f"S{r}"].value = "=" + hn_cell("K", r)
    tm = f'TRIM(M{r}&"")'
    dp[f"T{r}"].value = f'=IF({tm}="","",IF(ISNUMBER(MATCH({tm},코드표!$C$2:$C$11,0)),{tm},"⚠"&{tm}))'
    tl = f'TRIM(L{r}&"")'
    dp[f"U{r}"].value = f'=IF({tl}="","",IFERROR(INDEX(코드표!$G$2:$G$101,MATCH({tl},코드표!$F$2:$F$101,0)),"⚠"&{tl}))'
    dp[f"V{r}"].value = f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r})),I{r}-H{r}+1,"")'
    dp[f"W{r}"].value = (f'=IF(ISNUMBER(H{r}),MAX(0,MIN(IF(ISNUMBER(I{r}),I{r},{EOM_}),{EOM_})-MAX(H{r},{BOM})+1),0)')
    dp[f"X{r}"].value = (f'=IF(OR(NOT(ISNUMBER(H{r})),Q{r}=""),"",IF(TODAY()<H{r},"파견예정",'
                         f'IF(AND(ISNUMBER(I{r}),TODAY()>I{r}),"파견종료","파견중")))')
    dp[f"Y{r}"].value = f'=IF(OR(Q{r}="",X{r}=""),"",Q{r}&"|"&X{r})'
    dp[f"Z{r}"].value = f'=IF(AND(Q{r}<>"",W{r}>0),Q{r}&"|월","")'
    dp[f"AA{r}"].value = (
        f'=IF(COUNTA(B{r}:N{r})=0,"",IF(Q{r}="","등록번호 없음",IF(LEFT(R{r},1)="⚠","모병원명 미등록",'
        f'IF(S{r}="","파견병원 입력",IF(LEFT(S{r},1)="⚠","파견병원명 미등록",IF(LEFT(U{r},1)="⚠","파견과목 미등록",'
        f'IF(OR(T{r}="",LEFT(T{r},1)="⚠"),"파견근거 확인",IF(NOT(AND(ISNUMBER(H{r}),ISNUMBER(I{r}))),"날짜 형식/누락",'
        f'IF(I{r}<H{r},"종료일<시작일","")))))))))')
    note = f'TRIM(N{r}&" "&O{r})'
    ymd = lambda a, b, c: f'DATE(VALUE(MID({note},{a},4)),VALUE(MID({note},{b},2)),VALUE(MID({note},{c},2)))'
    dp[f"AB{r}"].value = f'=IF(ISNUMBER(SEARCH("다기관",{note})),IFERROR({ymd(1,6,9)},IF(ISNUMBER(H{r}),H{r},"")),"")'
    dp[f"AC{r}"].value = f'=IF(ISNUMBER(SEARCH("다기관",{note})),IFERROR({ymd(12,17,20)},IF(ISNUMBER(I{r}),I{r},"")),"")'
    dp[f"AD{r}"].value = f'=IF(AND(ISNUMBER(AB{r}),ISNUMBER(AC{r})),MAX(0,MIN(AC{r},{EOM_})-MAX(AB{r},{BOM})+1),0)'
    dp[f"AE{r}"].value = f'=IF(AND(Q{r}<>"",ISNUMBER(AB{r})),Q{r}&"|다기관","")'
    for col, _, _ in helpers:
        cell = dp[f"{col}{r}"]
        cell.font, cell.fill, cell.border, cell.alignment = f_base, fill_calc, border, center
    dp[f"W{r}"].number_format = "0"
    dp[f"AB{r}"].number_format = dp[f"AC{r}"].number_format = "yyyy-mm-dd"
    dp[f"AD{r}"].number_format = "0"

# 입력 검증 (직접 입력 시 오류 팝업)
def dp_list(rng_formula, target, what):
    dv = DataValidation(type="list", formula1=rng_formula, allow_blank=True, showErrorMessage=True,
                        errorStyle="stop", errorTitle="등록되지 않은 값",
                        error=f"{what} 목록에 없는 값입니다. 새로운 항목이면 '코드표' 시트에 먼저 추가한 뒤 입력하세요.")
    dp.add_data_validation(dv)
    dv.add(target)


dp_list("=코드표!$J$2:$J$101", f"B4:B{PLAN2_LAST}", "병원명")
dp_list("=코드표!$J$2:$J$101", f"K4:K{PLAN2_LAST}", "병원명")
dp_list("=코드표!$F$2:$F$101", f"L4:L{PLAN2_LAST}", "진료과")
dp_list("=코드표!$C$2:$C$11", f"M4:M{PLAN2_LAST}", "파견근거")
dv_ir = DataValidation(type="list", formula1='"I,R"', allow_blank=True, showErrorMessage=True,
                       errorTitle="수련종별", error="I(인턴) 또는 R(레지던트)로 입력하세요.")
dp.add_data_validation(dv_ir)
dv_ir.add(f"C4:C{PLAN2_LAST}")
dv_d = DataValidation(type="date", operator="greaterThan", formula1="1900-01-01", allow_blank=True,
                      showErrorMessage=True, errorTitle="날짜 형식", error="yyyy-mm-dd 형식의 날짜를 입력하세요.")
dp.add_data_validation(dv_d)
dv_d.add(f"H4:I{PLAN2_LAST}")
dv_e = DataValidation(type="custom", formula1='OR(I4="",H4="",I4>=H4)', allow_blank=True, showErrorMessage=True,
                      errorTitle="종료일 오류", error="종료일은 시작일 이후여야 합니다.")
dp.add_data_validation(dv_e)
dv_e.add(f"I4:I{PLAN2_LAST}")
dp.conditional_formatting.add(f"A4:O{PLAN2_LAST}", FormulaRule(formula=['$AA4<>""'], fill=PatternFill("solid", bgColor="FFC7CE")))
dp.conditional_formatting.add(f"R4:U{PLAN2_LAST}", FormulaRule(formula=['LEFT(R4,1)="⚠"'], font=Font(name=FONT, color="C00000", bold=True)))
dp.conditional_formatting.add(f"AA4:AA{PLAN2_LAST}", FormulaRule(formula=['$AA4<>""'], font=Font(name=FONT, color="C00000", bold=True)))

# ---------------------------------------------------------------- 전공의명단
ws = wb.create_sheet("전공의명단", 3)
# (머리글, 너비, 종류) 종류: calc=수식, lk=기본자료 조회, in=입력, key=사번
cols = [
    ("번호", 6, "calc"), ("사번", 11, "key"), ("이름", 10, "lk"), ("성별", 6, "lk"),                      # A-D
    ("생년월일", 12, "lk"), ("수련과목", 14, "lk"), ("연차", 11, "lk"), ("근무지", 10, "lk"),            # E-H
    ("수련개시일", 12, "lk"), ("전공의등록번호", 14, "lk"), ("의사면허번호", 13, "lk"), ("재직여부", 9, "lk"),  # I-L
    ("입사장소", 11, "in"), ("현재 근무병원", 13, "calc"), ("현재 진료과", 16, "calc"),                    # M-O
    ("파견근거", 10, "calc"), ("파견병원", 11, "calc"), ("파견진료과", 15, "calc"),                       # P-R
    ("파견시작일", 12, "calc"), ("파견종료일", 12, "calc"),                                              # S-T
    ("파견상태", 10, "calc"), ("파견기간(일)", 10, "calc"), ("종료까지(일)", 10, "calc"),                 # U-W
    ("급여기준월 파견일수", 11, "calc"), ("급여지급병원(기준월)", 14, "calc"), ("점검", 26, "calc"),      # X-Z
    ("비고", 20, "in"), ("기본자료행", 9, "calc"), ("인턴계획행", 9, "calc"), ("인턴 이번달", 18, "calc"),  # AA-AD
    ("소속병원", 11, "calc"), ("파견기록행", 9, "calc"), ("기준월 근무병원", 12, "calc"),                 # AE-AG
    ("서울 근무일(기준월)", 10, "calc"), ("구리 근무일(기준월)", 10, "calc"), ("기준월 대표파견행", 10, "calc"),  # AH-AJ
    ("다기관 기록행", 9, "calc"), ("다기관(비고 기재 기간)", 26, "calc"),                                  # AK-AL
]
NCOL = len(cols)
for i, (h, w, k) in enumerate(cols, 1):
    c = ws.cell(row=1, column=i, value=h)
    c.font, c.border, c.alignment = f_head, border, wrap_center
    c.fill = fill_head if (k in ("in", "key") or i == 1) else fill_head2
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 42

BR = lambda col: f"기본자료!${col}$2:${col}${BLAST}"
PR = lambda col: f"{PL}!${col}$3:${col}${PLAN_LAST}"
DR = lambda col: f"{PL2}!${col}$4:${col}${PLAN2_LAST}"
lookup = {3: "C", 4: "J", 5: "K", 6: "D", 7: "E", 8: "F", 9: "G", 10: "H", 11: "I", 12: "A"}
date_cols = {5, 9}


def lk_formula(c, r):
    col = lookup[c]
    raw = f"INDEX({BR(col)},$AB{r})"
    if c in date_cols:
        # YYYYMMDD 숫자 -> 날짜. 형식이 맞지 않으면(예: 19999999) 원문을 텍스트로 표시
        return (f'=IF($AB{r}="","",IF({raw}="","",IF(NOT(ISNUMBER({raw})),{raw},'
                f'IF({raw}<1000000,{raw},'
                f'IF(AND({raw}>=10000101,INT(MOD({raw},10000)/100)>=1,INT(MOD({raw},10000)/100)<=12,'
                f'MOD({raw},100)>=1,MOD({raw},100)<=31),'
                f'DATE(INT({raw}/10000),INT(MOD({raw},10000)/100),MOD({raw},100)),{raw}&"")))))')
    return f'=IF($AB{r}="","",IF({raw}="","",{raw}))'


# 가상 예시: 명단에는 사번과 입사장소만 입력 (파견 정보는 파견수련계획에서 자동 연결)
sample = [
    ("2261212", "서울병원", "예시 데이터"), ("2261213", "서울병원", "예시 데이터"),
    ("2261214", "서울병원", "예시 데이터"), ("2261215", "서울병원", "예시 데이터(종료됨)"),
    ("2261216", "구리병원", "예시 데이터"), ("222222", "서울병원", "예시 데이터(인턴)"),
    ("222223", "서울병원", "예시 데이터(인턴)"), ("222224", "구리병원", "예시 데이터(인턴)"),
]
MD = "DAY(EOMONTH(요약!$B$3,0))"
for r in range(2, LAST + 1):
    KEY = f'TRIM($J{r}&"")'
    ws.cell(row=r, column=1, value=f'=IF(B{r}="","",ROW()-1)')
    if r - 2 < len(sample):
        ws.cell(row=r, column=2, value=sample[r - 2][0])
        ws.cell(row=r, column=13, value=sample[r - 2][1])
        ws.cell(row=r, column=27, value=sample[r - 2][2])
    for c in lookup:
        ws.cell(row=r, column=c, value=lk_formula(c, r))
    home_base = f'IF($H{r}="","",IF(RIGHT($H{r},2)="병원",$H{r},$H{r}&"병원"))'
    ok_intern = f'AND($AD{r}<>"",LEFT($AD{r},1)<>"⚠")'
    ws.cell(row=r, column=14, value=(
        f'=IF($B{r}="","",IF(AND($U{r}="파견중",$Q{r}<>""),$Q{r},'
        f'IF({ok_intern},LEFT($AD{r},FIND(" ",$AD{r})-1),{home_base})))'))
    ws.cell(row=r, column=15, value=(
        f'=IF($B{r}="","",IF(AND($U{r}="파견중",$Q{r}<>""),IF($R{r}<>"",$R{r},"파견과 미입력"),'
        f'IF({ok_intern},MID($AD{r},FIND(" ",$AD{r})+1,60),IF($F{r}="","",IF($F{r}="인턴","수련교육부(인턴)",$F{r})))))'))
    # 파견 정보: 파견수련계획에서 (파견중 > 파견예정 > 파견종료 순으로 대표 기록 1건) 자동 연결
    ws.cell(row=r, column=16, value=(
        f'=IF($B{r}="","",IF($AK{r}<>"",IF(AND(TODAY()>=INDEX({DR("AB")},$AK{r}),'
        f'OR(NOT(ISNUMBER(INDEX({DR("AC")},$AK{r}))),TODAY()<=INDEX({DR("AC")},$AK{r}))),"다기관",'
        f'IF($AF{r}="","",INDEX({DR("T")},$AF{r}))),IF($AF{r}="","",INDEX({DR("T")},$AF{r}))))'))
    ws.cell(row=r, column=17, value=f'=IF($AF{r}="","",INDEX({DR("S")},$AF{r}))')
    ws.cell(row=r, column=18, value=f'=IF($AF{r}="","",INDEX({DR("U")},$AF{r}))')
    ws.cell(row=r, column=19, value=f'=IF($AF{r}="","",IF(INDEX({DR("H")},$AF{r})="","",INDEX({DR("H")},$AF{r})))')
    ws.cell(row=r, column=20, value=f'=IF($AF{r}="","",IF(INDEX({DR("I")},$AF{r})="","",INDEX({DR("I")},$AF{r})))')
    ws.cell(row=r, column=21, value=f'=IF($AF{r}="","",INDEX({DR("X")},$AF{r}))')
    ws.cell(row=r, column=22, value=f'=IF($AF{r}="","",INDEX({DR("V")},$AF{r}))')
    ws.cell(row=r, column=23, value=f'=IF(AND($U{r}="파견중",$T{r}<>""),$T{r}-TODAY(),"")')
    ws.cell(row=r, column=24, value=f'=IF(OR($B{r}="",{KEY}=""),"",SUMPRODUCT(({DR("Q")}={KEY})*{DR("W")}))')
    # 급여지급병원(기준월): 서울·구리 사이는 기준월에 15일 이상 근무한 병원이 한꺼번에 지급.
    #  그 외 병원이 섞이면: 통합수련=파견병원, 모자/다기관=소속병원
    th = "요약!$E$3"
    ws.cell(row=r, column=25, value=(
        f'=IF($B{r}="","",IF($AG{r}="","",IF(IF({KEY}="",0,SUMPRODUCT(({DR("Q")}={KEY})*{DR("AD")}))>0,$AE{r},'
        f'IF(N($X{r})<=0,$AG{r},IF($AH{r}+$AI{r}={MD},'
        f'IF(AND($AH{r}>={th},$AI{r}<{th}),"서울병원",IF(AND($AI{r}>={th},$AH{r}<{th}),"구리병원",'
        f'"⚠15일 판정불가(서울 "&$AH{r}&"일/구리 "&$AI{r}&"일)")),'
        f'IF($AJ{r}="",$AE{r},IF(INDEX({DR("T")},$AJ{r})="통합수련",INDEX({DR("S")},$AJ{r}),$AE{r})))))))'))
    ws.cell(row=r, column=26, value=(
        f'=IF($B{r}="","",IF($AB{r}="","기본자료에 없는 사번",IF(COUNTIF($B$2:$B${LAST},$B{r})>1,"중복 사번",'
        f'IF($AE{r}="","소속(입사장소) 입력",'
        f'IF(IF({KEY}="",0,SUMPRODUCT(({DR("Q")}={KEY})*({DR("AA")}<>"")))>0,"파견수련계획 입력오류 확인",'
        f'IF(LEFT($Y{r},1)="⚠","서울·구리 15일 판정 불가(양쪽 모두 이상/미만)",'
        f'IF(IF($AC{r}="",0,N(INDEX({PR("T")},$AC{r})))>0,"인턴계획 입력오류 확인",'
        f'IF(AND($AC{r}<>"",$U{r}<>"파견중",$N{r}<>$AE{r}),"인턴 타지역 근무: 파견 여부 확인","정상"))))))))'))
    m = lambda x: f'IFERROR(MATCH({x},{PR("D")},0),IFERROR(MATCH({x}&"",{PR("D")},0),IFERROR(MATCH(VALUE({x}),{PR("D")},0),"")))'
    mb = f'IFERROR(MATCH($B{r},{BR("B")},0),IFERROR(MATCH($B{r}&"",{BR("B")},0),IFERROR(MATCH(VALUE($B{r}),{BR("B")},0),"")))'
    ws.cell(row=r, column=28, value=f'=IF($B{r}="","",{mb})')
    ws.cell(row=r, column=29, value=f'=IF($B{r}="","",{m(f"$B{r}")})')
    ws.cell(row=r, column=30, value=f'=IF(OR($AC{r}="",인턴현황!$G$3=""),"",INDEX({PL}!$M$3:$R${PLAN_LAST},$AC{r},인턴현황!$G$3)&"")')
    ws.cell(row=r, column=31, value=(
        f'=IF($B{r}="","",IF($M{r}<>"",$M{r},IF($AF{r}<>"",IF(LEFT(INDEX({DR("R")},$AF{r}),1)<>"⚠",INDEX({DR("R")},$AF{r}),{home_base}),{home_base})))'))
    ws.cell(row=r, column=32, value=(
        f'=IF(OR($B{r}="",{KEY}=""),"",IFERROR(MATCH({KEY}&"|파견중",{DR("Y")},0),'
        f'IFERROR(MATCH({KEY}&"|파견예정",{DR("Y")},0),IFERROR(MATCH({KEY}&"|파견종료",{DR("Y")},0),""))))'))
    imon = f'IF(OR($AC{r}="",인턴현황!$K$2=""),"",INDEX({PL}!$M$3:$R${PLAN_LAST},$AC{r},인턴현황!$K$2)&"")'
    ws.cell(row=r, column=33, value=(
        f'=IF($B{r}="","",IF(AND({imon}<>"",LEFT({imon},1)<>"⚠"),LEFT({imon},FIND(" ",{imon})-1),$AE{r}))'))
    for c, hosp in ((34, "서울병원"), (35, "구리병원")):
        ws.cell(row=r, column=c, value=(
            f'=IF($B{r}="","",IF({KEY}="",0,SUMPRODUCT(({DR("Q")}={KEY})*({DR("S")}="{hosp}")*{DR("W")}))'
            f'+IF($AG{r}="{hosp}",MAX(0,{MD}-N($X{r})),0))'))
    ws.cell(row=r, column=36, value=f'=IF(OR($B{r}="",{KEY}=""),"",IFERROR(MATCH({KEY}&"|월",{DR("Z")},0),""))')
    ws.cell(row=r, column=37, value=f'=IF(OR($B{r}="",{KEY}=""),"",IFERROR(MATCH({KEY}&"|다기관",{DR("AE")},0),""))')
    ws.cell(row=r, column=38, value=(
        f'=IF($AK{r}="","","다기관 "&TEXT(INDEX({DR("AB")},$AK{r}),"yyyy-mm-dd")&"~"&'
        f'IF(ISNUMBER(INDEX({DR("AC")},$AK{r})),TEXT(INDEX({DR("AC")},$AK{r}),"yyyy-mm-dd"),""))'))
    for c in range(1, NCOL + 1):
        cell = ws.cell(row=r, column=c)
        cell.border = border
        kind = cols[c - 1][2]
        cell.font = f_input if kind in ("key", "in") else f_base
        if kind in ("calc", "lk"):
            cell.fill = fill_calc
        if c not in (26, 27):
            cell.alignment = center
    for c in (5, 9, 19, 20):
        ws.cell(row=r, column=c).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=2).number_format = "@"
    for c in (22, 23, 24, 34, 35):
        ws.cell(row=r, column=c).number_format = "0"

tab = Table(displayName="전공의", ref=f"A1:{get_column_letter(NCOL)}{LAST}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
ws.add_table(tab)
ws.freeze_panes = "D2"

dv_m = DataValidation(type="list", formula1="=코드표!$A$2:$A$21", allow_blank=True, showErrorMessage=True,
                      errorStyle="stop", errorTitle="등록되지 않은 값",
                      error="병원 목록에 없는 값입니다. 새로운 병원이면 '코드표' 시트에 먼저 추가한 뒤 선택하세요.")
ws.add_data_validation(dv_m)
dv_m.add(f"M2:M{LAST}")

cf = ws.conditional_formatting
cf.add(f"B2:B{LAST}", FormulaRule(formula=['AND($B2<>"",$AB2="")'], fill=PatternFill("solid", bgColor="FF0000"), font=Font(name=FONT, color="FFFFFF", bold=True)))
cf.add(f"B2:B{LAST}", FormulaRule(formula=['AND($B2<>"",COUNTIF($B$2:$B$%d,$B2)>1)' % LAST], fill=PatternFill("solid", bgColor="FFC000")))
cf.add(f"E2:E{LAST}", FormulaRule(formula=['ISTEXT($E2)'], font=Font(name=FONT, color="FF0000", bold=True)))
cf.add(f"I2:I{LAST}", FormulaRule(formula=['ISTEXT($I2)'], font=Font(name=FONT, color="FF0000", bold=True)))
cf.add(f"Z2:Z{LAST}", FormulaRule(formula=['AND($Z2<>"",$Z2<>"정상")'], fill=PatternFill("solid", bgColor="FFC7CE"), font=Font(name=FONT, color="9C0006", bold=True)))
cf.add(f"U2:W{LAST}", FormulaRule(formula=['$U2="파견종료"'], fill=PatternFill("solid", bgColor="D9D9D9"), font=Font(name=FONT, color="7F7F7F")))
cf.add(f"A2:T{LAST}", FormulaRule(formula=[f'AND($U2="파견중",$W2<>"",$W2<={xref("요약!$E$2")})'], fill=PatternFill("solid", bgColor="F8CBAD")))
cf.add(f"U2:U{LAST}", FormulaRule(formula=['$U2="파견중"'], fill=PatternFill("solid", bgColor="C6E0B4")))
cf.add(f"U2:U{LAST}", FormulaRule(formula=['$U2="파견예정"'], fill=PatternFill("solid", bgColor="BDD7EE")))
cf.add(f"Y2:Y{LAST}", FormulaRule(formula=['LEFT($Y2,1)="⚠"'], fill=PatternFill("solid", bgColor="FF0000"), font=Font(name=FONT, color="FFFFFF", bold=True)))
cf.add(f"Y2:Y{LAST}", FormulaRule(formula=['AND($Y2<>"",$AE2<>"",$Y2<>$AE2)'], fill=PatternFill("solid", bgColor="FFFF00")))

# ---------------------------------------------------------------- 요약
sm = wb.create_sheet("요약", 4)
sm["A1"].value, sm["A1"].font = "파견 현황 요약", f_title
sm["A2"].value, sm["A2"].font = "기준일", f_bold
sm["B2"].value, sm["B2"].number_format, sm["B2"].font = "=TODAY()", "yyyy-mm-dd", f_base
sm["D2"].value, sm["D2"].font = "종료 임박 기준(일)", f_bold
sm["E2"].value, sm["E2"].font, sm["E2"].fill = 30, f_input, fill_key
sm["F2"].value, sm["F2"].font = "← 전공의명단 주황색 강조에도 반영", f_base
sm["A3"].value, sm["A3"].font = "급여 기준월(1일)", f_bold
sm["B3"].value, sm["B3"].number_format, sm["B3"].font, sm["B3"].fill = "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)", "yyyy-mm-dd", f_input, fill_key
sm["D3"].value, sm["D3"].font = "서울·구리 15일 기준(일)", f_bold
sm["E3"].value, sm["E3"].font, sm["E3"].fill = 15, f_input, fill_key
sm["F3"].value, sm["F3"].font = "← 급여 기준월은 다른 달의 1일로 바꿔 입력하면 그 달 기준으로 계산됩니다", f_base


def head(cell, text):
    sm[cell].value, sm[cell].font, sm[cell].fill, sm[cell].alignment, sm[cell].border = text, f_head, fill_head, wrap_center, border


def body(cell, value, bold=False):
    sm[cell].value, sm[cell].font, sm[cell].border = value, f_bold if bold else f_base, border
    sm[cell].alignment = center


N = "전공의명단"
rg = lambda col: f"{N}!${col}$2:${col}${LAST}"
head("A4", "파견상태"); head("B4", "인원")
for i, s_ in enumerate(["파견중", "파견예정", "파견종료"]):
    body(f"A{5+i}", s_)
    body(f"B{5+i}", f"=COUNTIF({rg('U')},A{5+i})")
body("A8", "명단 등록 인원", True)
body("B8", f"=COUNTA({rg('B')})", True)
body("A9", "종료 임박(파견중)")
body("B9", f'=COUNTIFS({rg("U")},"파견중",{rg("W")},"<="&$E$2)')

head("D4", "병원"); head("E4", "현재 근무"); head("F4", "급여 지급(기준월)"); head("G4", "파견 수용(파견중)")
for i in range(20):
    r = 5 + i
    body(f"D{r}", f'=IF(코드표!A{2+i}="","",코드표!A{2+i})')
    body(f"E{r}", f'=IF(D{r}="","",COUNTIF({rg("N")},D{r}))')
    body(f"F{r}", f'=IF(D{r}="","",COUNTIF({rg("Y")},D{r}))')
    body(f"G{r}", f'=IF(D{r}="","",COUNTIFS({rg("Q")},D{r},{rg("U")},"파견중"))')

head("I4", "현재 진료과"); head("J4", "인원")
for i in range(40):
    r = 5 + i
    body(f"I{r}", f'=IF(코드표!B{2+i}="","",코드표!B{2+i})')
    body(f"J{r}", f'=IF(I{r}="","",COUNTIF({rg("O")},I{r}))')
for i in range(2):  # 인턴 전용 구분
    r = 45 + i
    body(f"I{r}", f'=IF(코드표!D{2+i}="","",코드표!D{2+i}&" (인턴)")')
    body(f"J{r}", f'=IF(코드표!D{2+i}="","",COUNTIF({rg("O")},코드표!D{2+i}))')

head("A12", "파견근거(종류)"); head("B12", "인원")
for i in range(8):
    r = 13 + i
    body(f"A{r}", f'=IF(코드표!C{2+i}="","",코드표!C{2+i})')
    body(f"B{r}", f'=IF(A{r}="","",COUNTIF({rg("P")},A{r}))')

head("A22", "데이터 점검"); head("B22", "건수")
checks = [
    ("기본자료 인원", f"=COUNTA({BR('B')})"),
    ("기본자료 재직 인원", f'=COUNTIF({BR("A")},"재직")'),
    ("기본자료에 없는 사번", f'=SUMPRODUCT(({rg("B")}<>"")*({rg("AB")}=""))'),
    ("명단 중복 사번", f'=SUMPRODUCT(({rg("B")}<>"")*(COUNTIF({rg("B")},{rg("B")})>1))'),
    ("날짜 형식 오류(생년월일/수련개시일)", f'=SUMPRODUCT(({rg("E")}<>"")*1)-COUNT({rg("E")})+SUMPRODUCT(({rg("I")}<>"")*1)-COUNT({rg("I")})'),
    ("인턴계획 입력 오류(병원/진료과)", f"=SUM({PL}!$T$3:$T${PLAN_LAST})"),
    ("파견수련계획 입력 오류", f'=SUMPRODUCT(--(LEN({PL2}!$AA$4:$AA${PLAN2_LAST})>0))'),
    ("명단에 없는 파견 기록(등록번호)", f'=SUMPRODUCT(({PL2}!$Q$4:$Q${PLAN2_LAST}<>"")*ISNA(MATCH({PL2}!$Q$4:$Q${PLAN2_LAST},{rg("J")},0)))'),
    ("서울·구리 15일 판정 불가(양쪽 모두 이상/미만)", f'=COUNTIF({rg("Y")},"⚠*")'),
    ("점검 필요 행(전공의명단 '점검' 열)", f'=SUMPRODUCT(({rg("B")}<>"")*({rg("Z")}<>"정상"))'),
]
for i, (label, fml) in enumerate(checks):
    body(f"A{23+i}", label)
    body(f"B{23+i}", fml)
sm.conditional_formatting.add("B25:B32", FormulaRule(formula=["B25>0"], fill=PatternFill("solid", bgColor="FFC7CE")))
for col, w in zip("ABCDEFGHIJ", [34, 10, 3, 14, 12, 14, 14, 3, 20, 10]):
    sm.column_dimensions[col].width = w
sm.row_dimensions[4].height = 30

# ---------------------------------------------------------------- 인턴현황 (조회·정렬용)
iv = wb.create_sheet("인턴현황", 5)
iv["A1"].value, iv["A1"].font = "인턴 근무 현황", f_title
iv["A2"].value, iv["A2"].font = "기준일", f_bold
iv["B2"].value, iv["B2"].number_format, iv["B2"].font = "=TODAY()", "yyyy-mm-dd", f_base
iv["A3"].value, iv["A3"].font = "계획 월 번호(자동)", f_note
iv["F3"].value, iv["F3"].font = "이번/다음 달 위치→", f_note
iv["J2"].value, iv["J2"].font = "급여 기준월 위치→", f_note
iv["K2"].value = '=IFERROR(MATCH(MONTH(요약!$B$3),$K$3:$P$3,0),"")'
iv["K2"].font, iv["K2"].alignment = f_note, center
iv["G3"].value = '=IFERROR(MATCH(MONTH(TODAY()),$K$3:$P$3,0),"")'
iv["I3"].value = '=IFERROR(MATCH(MONTH(EDATE(TODAY(),1)),$K$3:$P$3,0),"")'
for c in ("G3", "I3"):
    iv[c].font, iv[c].alignment = f_note, center
ih = {1: ("계획행", 7), 2: ("순번", 6), 3: ("이름", 10), 4: ("사번", 10), 5: ("기본자료 확인", 11), 6: ("입사일", 12),
      7: ('="이번 달("&IF($G$3="","계획 외",MONTH(TODAY())&"월")&") 병원"', 15), 8: ("이번 달 진료과", 16),
      9: ('="다음 달("&IF($I$3="","계획 외",MONTH(EDATE(TODAY(),1))&"월")&") 병원"', 15), 10: ("다음 달 진료과", 16),
      17: ("입력오류(건)", 9)}
for k in range(6):
    ih[11 + k] = (f'=TRIM({PL}!{MONTH_SRC[k]}$2&"")', 18)
for c, (h, w) in ih.items():
    cell = iv.cell(row=5, column=c, value=h)
    cell.font, cell.fill, cell.border, cell.alignment = f_head, fill_head, border, wrap_center
    iv.column_dimensions[get_column_letter(c)].width = w
iv.row_dimensions[5].height = 30
for k in range(6):
    col = get_column_letter(11 + k)
    iv[f"{col}3"].value = f'=IFERROR(VALUE(SUBSTITUTE(TRIM({col}5),"월","")),"")'
    iv[f"{col}3"].font, iv[f"{col}3"].alignment = f_note, center

IV_FIRST = 6
IV_LAST = IV_FIRST + INTERN_CAP - 1
for i in range(INTERN_CAP):
    r = IV_FIRST + i
    iv.cell(row=r, column=1, value=i + 1)  # 계획행 번호(상수): 정렬해도 같은 행끼리 유지
    iv.cell(row=r, column=3, value=f'=IFERROR(TRIM(INDEX({PR("C")},$A{r})&""),"")')
    iv.cell(row=r, column=2, value=f'=IF($C{r}="","",INDEX({PR("B")},$A{r}))')
    iv.cell(row=r, column=4, value=f'=IF($C{r}="","",INDEX({PR("D")},$A{r}))')
    iv.cell(row=r, column=5, value=(
        f'=IF($C{r}="","",IF(IFERROR(MATCH($D{r},{BR("B")},0),IFERROR(MATCH($D{r}&"",{BR("B")},0),'
        f'IFERROR(MATCH(VALUE($D{r}),{BR("B")},0),0)))>0,"확인","미등록"))'))
    iv.cell(row=r, column=6, value=f'=IF($C{r}="","",INDEX({PR("S")},$A{r}))')
    cur = f'INDEX({PL}!$M$3:$R${PLAN_LAST},$A{r},$G$3)&""'
    nxt = f'INDEX({PL}!$M$3:$R${PLAN_LAST},$A{r},$I$3)&""'
    for c, idx, expr, part in ((7, "G", cur, "h"), (8, "G", cur, "d"), (9, "I", nxt, "h"), (10, "I", nxt, "d")):
        if part == "h":
            res = f'IF({expr}="","",IF(LEFT({expr},1)="⚠",{expr},LEFT({expr},FIND(" ",{expr})-1)))'
        else:
            res = f'IF({expr}="","",IF(LEFT({expr},1)="⚠","",MID({expr},FIND(" ",{expr})+1,60)))'
        iv.cell(row=r, column=c, value=f'=IF(OR($C{r}="",${idx}$3=""),"",{res})')
    for k in range(6):
        iv.cell(row=r, column=11 + k, value=f'=IF($C{r}="","",INDEX({PR(MONTH_NORM[k])},$A{r})&"")')
    iv.cell(row=r, column=17, value=f'=IF($C{r}="","",INDEX({PR("T")},$A{r}))')
    for c in range(1, 18):
        cell = iv.cell(row=r, column=c)
        cell.border, cell.alignment = border, center
        cell.font = f_note if c == 1 else f_base
    iv.cell(row=r, column=6).number_format = "yyyy-mm-dd"
iv.freeze_panes = "D6"
iv.auto_filter.ref = f"A5:Q{IV_LAST}"

cf = iv.conditional_formatting
cf.add(f"E{IV_FIRST}:E{IV_LAST}", FormulaRule(formula=[f'$E{IV_FIRST}="미등록"'], fill=PatternFill("solid", bgColor="FFC7CE")))
cf.add(f"K{IV_FIRST}:P{IV_LAST}", FormulaRule(formula=[f'LEFT(K{IV_FIRST},1)="⚠"'], fill=PatternFill("solid", bgColor="FFC7CE"), font=Font(name=FONT, color="C00000", bold=True)))
cf.add(f"K5:P{IV_LAST}", FormulaRule(formula=['K$3=MONTH(TODAY())'], fill=PatternFill("solid", bgColor="FFF2CC")))
cf.add(f"G{IV_FIRST}:H{IV_LAST}", FormulaRule(formula=['$G$3<>""'], fill=PatternFill("solid", bgColor="FFF2CC")))

# 월별 병원별 인원 매트릭스
iv["S4"].value, iv["S4"].font = "월별 병원별 인턴 인원", f_bold
iv.column_dimensions["R"].width = 3
iv.cell(row=5, column=19, value="병원")
for k in range(6):
    iv.cell(row=5, column=20 + k, value=f"={get_column_letter(11 + k)}5")
for c in range(19, 26):
    cell = iv.cell(row=5, column=c)
    cell.font, cell.fill, cell.border, cell.alignment = f_head, fill_head, border, wrap_center
    iv.column_dimensions[get_column_letter(c)].width = 14 if c == 19 else 10
for i in range(20):
    r = IV_FIRST + i
    iv.cell(row=r, column=19, value=f'=IF(코드표!A{2+i}="","",코드표!A{2+i})')
    for k in range(6):
        col = get_column_letter(11 + k)
        iv.cell(row=r, column=20 + k, value=f'=IF($S{r}="","",COUNTIF({col}${IV_FIRST}:{col}${IV_LAST},$S{r}&" *"))')
    for c in range(19, 26):
        cell = iv.cell(row=r, column=c)
        cell.font, cell.border, cell.alignment = f_base, border, center
r = IV_FIRST + 20
iv.cell(row=r, column=19, value="⚠ 오류(인식 불가)")
iv.cell(row=r + 1, column=19, value="합계")
for k in range(6):
    col = get_column_letter(11 + k)
    oc = get_column_letter(20 + k)
    iv.cell(row=r, column=20 + k, value=f'=COUNTIF({col}${IV_FIRST}:{col}${IV_LAST},"⚠*")')
    iv.cell(row=r + 1, column=20 + k, value=f"=SUM({oc}{IV_FIRST}:{oc}{r})")
for rr in (r, r + 1):
    for c in range(19, 26):
        cell = iv.cell(row=rr, column=c)
        cell.font, cell.border, cell.alignment = f_bold, border, center
iv.cell(row=r + 2, column=19, value="※ 합계는 계획표에 입력된 인턴 수와 같아야 합니다.").font = f_note

# ---------------------------------------------------------------- 사용 안내
gd = wb.create_sheet("사용안내", 0)
lines = [
    ("전공의·인턴 파견 모니터링 사용 안내", f_title),
    ("", f_base),
    ("[전체 흐름]", f_bold),
    ("1) 기본자료  2) 인턴근무계획  3) 파견수련계획 — 세 시트에 각 양식 그대로 붙여넣기  4) 전공의명단에 사번·입사장소만 입력  5) 요약·인턴현황에서 확인", f_base),
    ("", f_base),
    ("[기본자료 업로드]", f_bold),
    (f"• 기본자료 시트 A1:L1 머리글 순서(현재근무 여부·사번·이름·수련과목·연차·근무지·수련개시일·전공의등록번호·의사면허번호·성별·생년월일·비고)로 A2부터 붙여넣으세요. 최대 {BASE_CAP}명. 열 순서는 유지해야 합니다.", f_base),
    ("• 날짜(수련개시일·생년월일)는 20260301 형식 숫자도 자동으로 날짜로 변환됩니다. 변환이 안 되는 값(예: 19999999)은 빨간 글씨 원문으로 표시되고 요약의 '날짜 형식 오류'에 집계됩니다.", f_base),
    ("• 기본자료 2~9행은 가상 예시입니다. 삭제하고 실제 자료를 붙여넣으세요. 인턴도 올리고 수련과목을 '인턴'으로 적으면 현재 진료과가 수련교육부(인턴)로 표시됩니다.", f_base),
    ("", f_base),
    ("[인턴근무계획 업로드]", f_bold),
    (f"• 인턴근무계획 시트의 업로드 양식(B1 제목, B2:K2 머리글, 3행부터 자료) 그대로 붙여넣으세요. 최대 {PLAN_LAST - 2}명. 오른쪽 초록 열(M~T)은 자동 정규화 결과이므로 지우지 마세요.", f_base),
    ("• 월 칸은 '지역 진료과'(예: 서울 내과, 서울 산부, 제주 에스중앙 내과)로 읽습니다. 첫 단어=지역(병원 표기표로 정식 병원 변환), 마지막 단어=진료과(약칭표로 정식 명칭 변환). 사이 단어는 무시합니다.", f_base),
    ("• 통합(과 미지정)과 핵방/핵/방(핵의학과+방사선종양학과)은 인턴 전용 구분으로 인정됩니다.", f_base),
    ("• 인턴근무계획에만 있는 내용은 해당 월 1일~말일 근무로 간주합니다(기본값). 비고의 2026.03은 2026년 3월 1일 입사로 해석해 인턴현황 '입사일'에 표시합니다.", f_base),
    ("", f_base),
    ("[파견수련계획 업로드]", f_bold),
    (f"• 파견수련계획 시트의 양식(1~3행 병합 머리글, 4행부터 자료: 연번·모병원명·수련종별·수련과목·연차·성명·전공의등록번호·파견시작일·파견종료일·파견일수·파견병원·파견과목·파견근거·비고)대로 붙여넣으세요. 최대 {PLAN2_LAST - 3}건.", f_base),
    ("• 한 사람이 여러 번 파견되면 행을 여러 개 입력합니다. 전공의명단과는 '전공의등록번호'로 연결됩니다(기본자료의 등록번호와 같아야 함).", f_base),
    ("• 모병원명·파견병원은 코드표의 '병원 표기' 표로 정식 병원(서울병원 등)에 연결됩니다. 표에 없는 기관명(인천·창원·오산 등)은 코드표 J·K열에 추가하세요.", f_base),
    ("• 파견근거는 모자 / 통합수련 / 다기관 입니다. 비고(N열, 또는 O열)에 '다기관'이 들어 있으면 다기관 파견으로 인식합니다. 비고가 '2026.09.01-2026.09.30 다기관 파견'처럼 기간으로 시작하면 그 기간을, 아니면 해당 행의 파견 기간을 다기관 기간으로 봅니다.", f_base),
    ("• 다기관 기간에는 전공의명단 '파견근거'에 다기관으로 표시되고, 오른쪽 끝 '다기관(비고 기재 기간)' 열에 기간이 보입니다. 급여 기준월에 다기관 기간이 걸리면 급여는 소속병원이 지급하는 것으로 계산합니다(가정).", f_base),
    ("• 코드표에 없는 병원·진료과·파견근거를 직접 입력하면 오류 팝업이 뜹니다. 붙여넣기는 엑셀 특성상 팝업이 뜨지 않으므로, 대신 해당 행이 빨갛게 표시되고 오른쪽 '입력 오류' 열과 요약에 집계됩니다.", f_base),
    ("", f_base),
    ("[전공의명단]", f_bold),
    ("• 입력은 사번, 입사장소, 비고뿐입니다. 인적사항은 기본자료에서, 파견 정보(파견근거·병원·과·기간·상태)는 파견수련계획에서 자동 연결됩니다(파견중 > 파견예정 > 파견종료 순으로 대표 1건 표시).", f_base),
    ("• 입사장소 = 채용된 병원(구리병원에서 채용해 근무하면 구리병원). 비워 두면 파견수련계획의 모병원명, 그것도 없으면 기본자료 근무지를 사용합니다.", f_base),
    ("• 현재 근무병원/진료과: ①파견중이면 파견병원/파견과목 ②인턴이면 이번 달 근무계획 ③그 외는 기본자료의 근무지/수련과목.", f_base),
    ("", f_base),
    ("[급여지급병원(기준월) 규칙 — 전공의명단 Y열]", f_bold),
    ("• 서울병원·구리병원 사이: 소속과 무관하게 급여 기준월에 15일 이상 근무한 병원이 그 달 전체를 한꺼번에 지급합니다. (예: 구리 10일·서울 20일이면 구리 10일치까지 서울병원에서 지급)", f_base),
    ("• 근무일 계산: 파견 기간은 파견병원, 나머지 날은 소속병원(인턴은 해당 월 근무계획 병원)에서 근무한 것으로 봅니다. 양쪽 모두 15일 이상이거나 모두 미만이면 판정할 수 없으므로 급여지급병원 칸에 빨간 ⚠ 오류가 표시되고, 점검 열과 요약에 집계됩니다.", f_base),
    ("• 서울·구리 외 병원이 섞인 경우(가정): 통합수련=파견병원이 지급, 모자·다기관=소속병원이 지급. 이 부분은 확정 규칙이 아니므로 실제 규정에 맞는지 확인하세요.", f_base),
    ("• 파견이 없는 달은 그 달 근무병원(소속 또는 인턴 근무계획 병원)이 지급합니다.", f_base),
    ("• 급여 기준월은 요약 시트 B3(기본값=이번 달 1일)입니다. 다른 달 급여는 그 달 1일 날짜로 바꾸세요. 15일 기준은 요약 E3에서 바꿀 수 있습니다. 급여지급병원 칸이 노란색이면 소속과 다른 병원에서 지급되는 경우입니다.", f_base),
    ("", f_base),
    ("[점검 열]", f_bold),
    ("• 전공의명단 '점검' 열이 '정상'이 아니면 빨간색으로 표시됩니다 (기본자료에 없는 사번, 중복, 소속 누락, 파견수련계획/인턴계획 입력오류, 인턴 타지역 근무인데 파견 입력 없음 등).", f_base),
    ("", f_base),
    ("[정렬·필터]", f_bold),
    ("• 전공의명단·인턴현황 머리글의 ▼ 버튼으로 정렬/필터합니다 (예: 파견상태=파견중, 종료까지(일) 오름차순 → 종료 임박순).", f_base),
    ("• 주황색 행: 파견중이면서 종료까지 요약 시트의 기준일수(기본 30일) 이내.  회색: 파견종료.", f_base),
    ("", f_base),
    ("[행 늘리기]", f_bold),
    (f"• 전공의명단은 {CAPACITY}행까지 수식이 채워져 있습니다. 더 필요하면 표의 마지막 행 '위'에 행을 삽입하고 바로 위 행을 복사해 붙여넣으세요. 마지막 행 아래에 붙이면 요약 등이 참조하는 범위에 포함되지 않습니다. (인턴근무계획·파견수련계획·인턴현황도 같은 방법)", f_base),
    ("", f_base),
    ("[참고]", f_bold),
    ("• 개인정보(생년월일·면허번호 등)가 포함되므로 파일 암호 설정과 접근 권한 관리를 권장합니다.", f_base),
]
for i, (t, f) in enumerate(lines, 1):
    gd.cell(row=i, column=1, value=t).font = f
gd.column_dimensions["A"].width = 175
gd.sheet_view.showGridLines = False

wb.move_sheet("코드표", offset=0)
wb.save(OUT)
print("saved", OUT, wb.sheetnames)
