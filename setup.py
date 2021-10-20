import pandas as pd  # 데이터프레임 생성용 패키지
import functions as func  # functions.py 파일의 함수 사용
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# 1. 필수 파일 존재 여부 확인
existence_number = func.tf_exist_all_files()
if existence_number == 4:
    quit(0)

# 2. 관련 변수 설정
df_course = func.summarize_course(existence_number)
student_number, df_student = func.summarize_student_information(existence_number)
df_elective = func.summarize_elective_course()

# 3. 강좌개설정보에 수강한 강좌 반영
"""
코드쉐어 과목 처리 방법
1. 강좌개설정보(df_course)에서 학생 수강 과목(df_student) 중 하나와 과목코드[전공분야코드, 난이도, 일련번호]가 일치하는 과목을 추출함.
2. 강좌개설정보에서 다음 조건을 만족하는 과목을 추가로 추출함.
    2-1. 1에서 추출한 과목과 [교과목명]이 같은 과목(자신 포함).
3. 2에서 만족하는 과목은 수강 횟수에 1을 더함.
"""
# 3-1. 강좌개설정보 변수(df_course)에 수강 횟수 정보 추가
df_course['수강횟수'] = 0

for row in range(len(df_student)):
    # 3-2. "코드쉐어 과목 처리 방법" 중 1번 항목 시행
    df_infected = df_course[df_course['전공분야코드'] == df_student.iloc[row]['전공분야코드']]
    df_infected = df_infected[df_infected['난이도'] == df_student.iloc[row]['난이도']]
    df_infected = df_infected[df_infected['일련번호'] == df_student.iloc[row]['일련번호']]

    # 3-3. "코드쉐어 과목 처리 방법" 중 2번 항목 시행
    df_total_infected = pd.DataFrame()
    for inf_row in range(len(df_infected)):
        # [교과목명]이 같은 과목
        df_add_infected1 = df_course[df_course['교과목명'] == df_infected.iloc[inf_row]['교과목명']]
        df_total_infected = pd.concat([df_total_infected, df_add_infected1])
    # 중복되는 과목을 하나로 제거
    df_total_infected = df_total_infected.drop_duplicates()

    # 3-4. "코드쉐어 과목 처리 방법" 중 3번 항목 시행
    list_index_total_infected = df_total_infected.index.values
    for i in list_index_total_infected:
        df_course.loc[df_course.index == i, "수강횟수"] = df_course.loc[df_course.index == i, "수강횟수"] + 1

# 과목 수강 횟수 확인용
# print(df_course[df_course["수강횟수"] != 0].sort_values(by=['교과목명'], axis=0))


# 4. 수강횟수가 반영된 강좌개설정보를 전공분야코드별로 엑셀에 저장
# 4-1. 전공분야코드를 최초개설년도 및 학기 순으로 정렬
df_major_code = df_course.groupby(["전공분야코드"], as_index=False)[["최초개설년도", "최초개설학기"]].min()
df_major_code = df_major_code.sort_values(by=['최초개설년도', '최초개설학기', '전공분야코드'])

# 4-2. 전공분야코드 중복 제거
list_major_code = list(dict.fromkeys(df_major_code["전공분야코드"].values.tolist()))
# print(list_major_code)

# 4-3. 전공분야코드별로 엑셀 파일(template.xlsx)에 개설강좌정보 입력
filename = 'template'
template = openpyxl.load_workbook("./"+filename+".xlsx")
sheet = template["개설과목정보"]

# 각 셀의 테두리 스타일
THIN_BORDER = Border(Side('thin'), Side('thin'), Side('thin'), Side('thin'))  # 좌우상하 순서

# 입력 시 개설강좌정보 최좌상단 셀 위치
start_row = 5
start_col = 2

for major in list_major_code:
    df_major_course = df_course[df_course["전공분야코드"] == major]  # 전공분야코드에 따라 선별한 개설강좌정보
    # 추후 개설강좌정보 업데이트 시 셀 순서가 뒤바뀌지 않도록 다음 조건에 따라 정렬
    df_major_course = df_major_course.sort_values(by=['최초개설년도', '최초개설학기', '난이도', '일련번호'])
    
    # 정보를 각 셀에 입력
    small_row = start_row
    for maj_row in dataframe_to_rows(df_major_course, index=False, header=True):
        small_col = start_col
        for value in maj_row:
            sheet.cell(row=small_row, column=small_col).value = value
            if small_row != start_row:
                sheet.cell(row=small_row, column=small_col).border = THIN_BORDER
            small_col += 1
        small_row += 1
    start_col += len(df_major_course.columns) + 1
    # print(df_major_course)

# 4-4. 엑셀 파일에 학번 및 최초개설학기 설명 입력
sheet.cell(row=2, column=6).value = "최초 개설\n학기 설명"
sheet.cell(row=2, column=7).value = "1: 1학기, 2: 여름학기\n3: 2학기, 4: 겨울학기"

# 4-5. 서식 및 디자인 설정
# 선 디자인
LEFT_BORDER = Border(Side('thick'), Side('thin'), Side('thick'), Side('thick'))  # 좌우상하 순서
RIGHT_BORDER = Border(Side('thin'), Side('thick'), Side('thick'), Side('thick'))
# 셀 색상
LIGHT = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type='solid')
DARK = PatternFill(start_color="31869B", end_color="31869B", fill_type='solid')

# 4-5-1. 행 높이, 열 너비 설정 및 틀 고정
# 행 높이
sheet.row_dimensions[1].height = 8
sheet.row_dimensions[2].height = 40
sheet.row_dimensions[3].height = 8
sheet.row_dimensions[4].height = 20
sheet.row_dimensions[5].height = 20

# 열 너비
for column in range(len(list_major_code)):
    if column == 0:
        sheet.column_dimensions[get_column_letter(9 * column + 1)].width = 1   # 빈 칸
    else:
        sheet.column_dimensions[get_column_letter(9 * column + 1)].width = 3  # 빈 칸
    sheet.column_dimensions[get_column_letter(9 * column + 2)].width = 15  # 최초개설년도
    sheet.column_dimensions[get_column_letter(9 * column + 3)].width = 15  # 최초개설학기
    sheet.column_dimensions[get_column_letter(9 * column + 4)].width = 12  # 전공분야코드
    sheet.column_dimensions[get_column_letter(9 * column + 5)].width = 9   # 난이도
    sheet.column_dimensions[get_column_letter(9 * column + 6)].width = 9   # 일련번호
    sheet.column_dimensions[get_column_letter(9 * column + 7)].width = 35  # 교과목명
    sheet.column_dimensions[get_column_letter(9 * column + 8)].width = 9   # 학점
    sheet.column_dimensions[get_column_letter(9 * column + 9)].width = 9   # 수강횟수

# 틀 고정
sheet.freeze_panes = "A6"


# 4-5-2. 학번 부분[B2:C2] 및 최초개설학기 설명[F2:G2] 디자인
# 선 디자인 반영
sheet.cell(row=2, column=2).border = LEFT_BORDER
sheet.cell(row=2, column=3).border = RIGHT_BORDER
sheet.cell(row=2, column=6).border = LEFT_BORDER
sheet.cell(row=2, column=7).border = RIGHT_BORDER
# 셀 색상 및 글꼴 굵기 반영(왼쪽 부분)
sheet.cell(row=2, column=2).fill = LIGHT
sheet.cell(row=2, column=2).font = Font(bold=True)
sheet.cell(row=2, column=6).fill = LIGHT
sheet.cell(row=2, column=6).font = Font(bold=True)
# 글자 서식 반영
sheet.cell(row=2, column=2).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
sheet.cell(row=2, column=3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
sheet.cell(row=2, column=6).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
sheet.cell(row=2, column=7).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# 4-5-3. 강좌개설정보 부분 디자인
for column in range(len(list_major_code)):
    # 셀 병합
    sheet.merge_cells(start_row=4, start_column=9 * column + 2, end_row=4,  end_column=9 * column + 9)
    # 색상 설정 및 글자 서식 반영
    sheet.cell(row=4, column=9 * column + 2).fill = LIGHT
    sheet.cell(row=4, column=9 * column + 2).alignment = Alignment(horizontal='center', vertical='center')
    sheet.cell(row=4, column=9 * column + 2).font = Font(bold=True)
    for i in range(2, 10):
        sheet.cell(row=5, column=9 * column + i).fill = DARK
        sheet.cell(row=5, column=9 * column + i).alignment = Alignment(horizontal='center', vertical='center')
        sheet.cell(row=5, column=9 * column + i).font = Font(bold=True, color='ffffff')

# 4-6. 각 전공분야코드별 설명 추가
# 전공분야코드별 설명
df_major_explain = pd.DataFrame({"전공분야코드": ["CC", "EC", "MS", "ME", "EN", "LS",
                                            "PH", "CH", "NA", "MD", "ET", "CT",
                                            "RT", "FE", "EP", "AI", "MI", "IC",
                                            "UC", "GS", "PS", "BS", "MC", "MA",
                                            "EV", "MM", "IR", "LH", "PP", "EB",
                                            "SS", "MB", "CM"
                                            # 이후 추가된 분야코드 존재 시 아랫줄에 추가
                                            ],
                                 "설명": [
                                        # 대학원 또는 학사·대학원 과정 전공 과목
                                        "대학원 : 공통과목", "학사·대학원 : 전기전자컴퓨터공학부",
                                        "대학원 : 신소재공학부", "대학원 : 기계공학부",
                                        "대학원 : 지구·환경공학부", "대학원 : 생명공학부",
                                        "대학원 : 물리·광과학과", "학사·대학원 : 화학과",
                                        "대학원 : 나노바이오재료전자공학과", "학사 : (부전공)의생명공학  |  대학원 : 의생명공학과",
                                        "학사 : (부전공)에너지  |  대학원 : 융합기술학제학부 - 에너지",
                                        "학사 : (부전공)문화기술  |  대학원:융합기술학제학부 - 문화기술",
                                        "대학원 : 융합기술학제학부 - 지능로봇", "학사 : (부전공)에너지  |  대학원 : 에너지융합대학원",
                                        "대학원 : (부전공)석사 창업", "대학원 : AI 대학원",
                                        "대학원 : (부전공)기술혁신", "대학원 : (舊)전기전자컴퓨터공학부",
                                        # 학사 과정 전용 전공 과목
                                        "학사 : 공통과목", "학사 : 기초교육학부",
                                        "학사 : 물리·광과학과", "학사 : 생명과학부",
                                        "학사 : 기계공학부", "학사 : 신소재공학부",
                                        "학사 : 지구·환경공학부", "학사 : (부전공)수학",
                                        "학사 : (부전공)지능로봇", "학사 : (부전공)인문사회 - 문화와 역사",
                                        "학사 : (부전공)인문사회 - 공공정책·법정치사회", "학사 : (부전공)인문사회 - 경제·경영",
                                        "학사 : (부전공)인문사회 - 과학기술과 사회", "학사 : (부전공)인문사회 - 마음과 행동",
                                        "학사 : (舊)화학과"
                                        # 이후 추가된 분야코드 존재 시 아랫줄에 추가
                                        ]})

# 혹여 추후 추가된 분야코드에서 중복되는 코드 발견 시 자동으로 삭제
df_major_explain = df_major_explain.drop_duplicates("전공분야코드")

# 엑셀 파일에 설명 추가
for column in range(len(list_major_code)):
    sheet.cell(row=4, column=9 * column + 2).value = \
        df_major_explain.loc[df_major_explain["전공분야코드"] == list_major_code[column], "설명"].values[0]

# 5. 교양 과목에 수강횟수를 반영하여 엑셀에 저장
sheet = template["교양과목정보"]

# 5-1. df_course 변수와 합쳐 교양 과목 변수(df_course)에 수강 횟수 정보 추가
df_elective = pd.merge(df_elective, df_course, how='inner', on=["전공분야코드", "난이도", "일련번호", "교과목명"])

# 5-2. 분류별로 엑셀 시트에 기입
# 입력 시 최좌상단 셀 위치
start_row = 5
start_col = 2

list_elective_code = list(dict.fromkeys(df_elective["분류"].values.tolist()))
for elect in list_elective_code:
    df_elective_course = df_elective[df_elective["분류"] == elect].drop(["분류"], axis=1)  # 교양과목 분류에 따라 선별한 개설강좌정보
    # 추후 개설강좌정보 업데이트 시 셀 순서가 뒤바뀌지 않도록 다음 조건에 따라 정렬
    df_elective_course = df_elective_course.sort_values(by=['최초개설년도', '최초개설학기', '난이도', '일련번호'])
    # 컬럼명 재배열
    df_elective_course = df_elective_course[["최초개설년도", "최초개설학기", "전공분야코드", "난이도", "일련번호", "교과목명", "학점", "수강횟수"]]

    # 정보를 각 셀에 입력
    small_row = start_row
    for elect_row in dataframe_to_rows(df_elective_course, index=False, header=True):
        small_col = start_col
        for value in elect_row:
            sheet.cell(row=small_row, column=small_col).value = value
            if small_row != start_row:
                sheet.cell(row=small_row, column=small_col).border = THIN_BORDER
            small_col += 1
        small_row += 1
    start_col += len(df_elective_course.columns) + 1
    # print(df_elective_course)

# 5-3. 엑셀 파일에 학번 및 최초개설학기 설명 입력
sheet.cell(row=2, column=6).value = "최초 개설\n학기 설명"
sheet.cell(row=2, column=7).value = "1: 1학기, 2: 여름학기\n3: 2학기, 4: 겨울학기"

# 5-4. 서식 및 디자인 설정
# 셀 색상
LIGHT = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type='solid')
DARK = PatternFill(start_color="548235", end_color="548235", fill_type='solid')

# 5-4-1. 행 높이, 열 너비 설정 및 틀 고정
# 행 높이
sheet.row_dimensions[1].height = 8
sheet.row_dimensions[2].height = 40
sheet.row_dimensions[3].height = 8
sheet.row_dimensions[4].height = 20
sheet.row_dimensions[5].height = 20

# 열 너비
for column in range(len(list_major_code)):
    if column == 0:
        sheet.column_dimensions[get_column_letter(9 * column + 1)].width = 1   # 빈 칸
    else:
        sheet.column_dimensions[get_column_letter(9 * column + 1)].width = 3  # 빈 칸
    sheet.column_dimensions[get_column_letter(9 * column + 2)].width = 15  # 최초개설년도
    sheet.column_dimensions[get_column_letter(9 * column + 3)].width = 15  # 최초개설학기
    sheet.column_dimensions[get_column_letter(9 * column + 4)].width = 12  # 전공분야코드
    sheet.column_dimensions[get_column_letter(9 * column + 5)].width = 9   # 난이도
    sheet.column_dimensions[get_column_letter(9 * column + 6)].width = 9   # 일련번호
    sheet.column_dimensions[get_column_letter(9 * column + 7)].width = 35  # 교과목명
    sheet.column_dimensions[get_column_letter(9 * column + 8)].width = 9   # 학점
    sheet.column_dimensions[get_column_letter(9 * column + 9)].width = 9   # 수강횟수

# 틀 고정
sheet.freeze_panes = "A6"


# 5-4-2. 학번 부분[B2:C2] 및 최초개설학기 설명[F2:G2] 디자인
# 선 디자인 반영
sheet.cell(row=2, column=2).border = LEFT_BORDER
sheet.cell(row=2, column=3).border = RIGHT_BORDER
sheet.cell(row=2, column=6).border = LEFT_BORDER
sheet.cell(row=2, column=7).border = RIGHT_BORDER
# 셀 색상 및 글꼴 굵기 반영(왼쪽 부분)
sheet.cell(row=2, column=2).fill = LIGHT
sheet.cell(row=2, column=2).font = Font(bold=True)
sheet.cell(row=2, column=6).fill = LIGHT
sheet.cell(row=2, column=6).font = Font(bold=True)
# 글자 서식 반영
sheet.cell(row=2, column=2).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
sheet.cell(row=2, column=3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
sheet.cell(row=2, column=6).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
sheet.cell(row=2, column=7).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# 5-4-3. 강좌개설정보 부분 디자인
for column in range(len(list_elective_code)):
    # 셀 병합
    sheet.merge_cells(start_row=4, start_column=9 * column + 2, end_row=4,  end_column=9 * column + 9)
    # 색상 설정 및 글자 서식 반영
    sheet.cell(row=4, column=9 * column + 2).fill = LIGHT
    sheet.cell(row=4, column=9 * column + 2).alignment = Alignment(horizontal='center', vertical='center')
    sheet.cell(row=4, column=9 * column + 2).font = Font(bold=True)
    for i in range(2, 10):
        sheet.cell(row=5, column=9 * column + i).fill = DARK
        sheet.cell(row=5, column=9 * column + i).alignment = Alignment(horizontal='center', vertical='center')
        sheet.cell(row=5, column=9 * column + i).font = Font(bold=True, color='ffffff')

# 5-5. 각 전공분야코드별 설명 추가
# 전공분야코드별 설명
df_elective_explain = pd.DataFrame({"분류": ["hus", "ppe", "gsc"],
                                    "설명": ["HUS : 문사철", "PPE : 철사과", "GSC : 일반선택"]})
# 엑셀 파일에 설명 추가
for column in range(len(list_elective_code)):
    sheet.cell(row=4, column=9 * column + 2).value = \
        df_elective_explain.loc[df_elective_explain["분류"] == list_elective_code[column], "설명"].values[0]

# 6. 성적 관련 정보를 엑셀에 저장
sheet = template["수강과목요약"]

# 6-1. 분류별로 엑셀 시트에 기입
# 입력 시 최좌상단 셀 위치
start_row = 5
start_col = 2

# 정보를 각 셀에 입력
small_row = start_row
for grade_row in dataframe_to_rows(df_student, index=False, header=True):
    small_col = start_col
    for value in grade_row:
        sheet.cell(row=small_row, column=small_col).value = value
        if small_row != start_row:
            sheet.cell(row=small_row, column=small_col).border = THIN_BORDER
        small_col += 1
    small_row += 1

# 6-2. 엑셀 파일에 학번 입력
sheet.cell(row=2, column=2).value = "학번"
sheet.cell(row=2, column=3).value = student_number

# 6-3. 서식 및 디자인 설정
# 셀 색상
LIGHT = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type='solid')
DARK = PatternFill(start_color="C65911", end_color="C65911", fill_type='solid')

# 6-3-1. 행 높이, 열 너비 설정 및 틀 고정
# 행 높이
sheet.row_dimensions[1].height = 8
sheet.row_dimensions[2].height = 40
sheet.row_dimensions[3].height = 8
sheet.row_dimensions[4].height = 20
sheet.row_dimensions[5].height = 20

# 열 너비
sheet.column_dimensions[get_column_letter(1)].width = 1   # 빈 칸
sheet.column_dimensions[get_column_letter(2)].width = 15  # 수강연도
sheet.column_dimensions[get_column_letter(3)].width = 15  # 수강학기
sheet.column_dimensions[get_column_letter(4)].width = 12  # 전공분야코드
sheet.column_dimensions[get_column_letter(5)].width = 9   # 난이도
sheet.column_dimensions[get_column_letter(6)].width = 9   # 일련번호
sheet.column_dimensions[get_column_letter(7)].width = 45  # 과목명
sheet.column_dimensions[get_column_letter(8)].width = 9   # 학점
sheet.column_dimensions[get_column_letter(9)].width = 9   # 수강횟수

# 틀 고정
sheet.freeze_panes = "A6"

# 6-3-2. 학번 부분[B2:C2] 디자인
# 선 디자인 반영
sheet.cell(row=2, column=2).border = LEFT_BORDER
sheet.cell(row=2, column=3).border = RIGHT_BORDER
# 셀 색상 및 글꼴 굵기 반영(왼쪽 부분)
sheet.cell(row=2, column=2).fill = LIGHT
sheet.cell(row=2, column=2).font = Font(bold=True)
# 글자 서식 반영
sheet.cell(row=2, column=2).alignment = Alignment(horizontal='center', vertical='center')
sheet.cell(row=2, column=3).alignment = Alignment(horizontal='center', vertical='center')

# 6-3-3. 강좌개설정보 부분 디자인
# 셀 병합
sheet.merge_cells(start_row=4, start_column=2, end_row=4,  end_column=9)
# 색상 설정 및 글자 서식 반영
sheet.cell(row=4, column=2).fill = LIGHT
sheet.cell(row=4, column=2).alignment = Alignment(horizontal='center', vertical='center')
sheet.cell(row=4, column=2).font = Font(bold=True)
for i in range(2, 10):
    sheet.cell(row=5, column=i).fill = DARK
    sheet.cell(row=5, column=i).alignment = Alignment(horizontal='center', vertical='center')
    sheet.cell(row=5, column=i).font = Font(bold=True, color='ffffff')

# 6-4. 엑셀 파일에 설명 추가
sheet.cell(row=4, column=2).value = "총 수강 과목"

# 7. 입력한 정보를 저장
template.save(filename="computation_result_kor.xlsx")
