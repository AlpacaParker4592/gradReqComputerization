import pandas as pd  # 데이터프레임 생성용 패키지
import functions as func  # functions.py 파일의 함수 사용

# 1. 필수 파일 존재 여부 확인
existence_number = func.tf_exist_all_files()
if existence_number == 4:
    quit(0)

# 2. 관련 변수 설정
df_course = func.summarize_course(existence_number)
student_number, df_student = func.summarize_student_information(existence_number)
df_elective = func.summarize_elective_course()
list_undergraduate_major_code = ["GS", "UC", "EC", "MA", "MC", "EV", "BS", "PS", "CH"]

# 엑셀 투입용 설명 데이터프레임
# 전공분야코드별 설명
df_major_explain = pd.read_excel("./data/code_explain.xlsx", sheet_name="major_code", keep_default_na=False)
# 혹여 추후 추가된 분야코드에서 중복되는 코드 발견 시 자동으로 삭제
df_major_explain = df_major_explain.drop_duplicates("전공분야코드")

# 교양 및 예체능 과목코드별 설명
df_elect_pna_res_explain = pd.read_excel("./data/code_explain.xlsx", sheet_name="elect_pna_res", keep_default_na=False)

# 3. 강좌개설정보에 수강한 강좌 반영
"""
코드쉐어 과목 처리 방법
1. 강좌개설정보(df_course)에서 학생 수강 과목(df_student) 중 하나와 과목코드[전공분야코드, 일련번호]가 일치하는 과목을 추출함.
2. 강좌개설정보에서 다음 조건을 만족하는 과목을 추가로 추출함.
    2-1. 1에서 추출한 과목과 [교과목명]이 같은 과목(자신 포함).
3. 2에서 만족하는 과목은 수강 횟수에 1을 더함.
4. 학생 수강 과목 내 모든 과목을 조회할 때까지 1부터 3 과정을 반복함.
"""
# 3-1. 강좌개설정보 변수(df_course)에 수강 횟수 정보 추가
df_course['수강횟수'] = 0
# 평점이 U 또는 F인 과목은 제외
df_student_not_u = df_student["평점"] != "U"
df_student_not_f = df_student["평점"] != "F"
df_student2 = df_student[df_student_not_u & df_student_not_f]

for row in range(len(df_student)):
    # 3-2. "코드쉐어 과목 처리 방법" 중 1번 항목 시행
    df_infected = df_course[df_course['전공분야코드'] == df_student2.iloc[row]['전공분야코드']]
    df_infected = df_infected[df_infected['일련번호'] == df_student2.iloc[row]['일련번호']]

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


# 4. 엑셀 시트에 넣을 데이터프레임 세분화(이수과목 시트 제외)
# 4-1. 교양-예체능-연구 과목 관련 데이터프레임 정의
# 4-1-1. 교양 과목 데이터프레임 정의
# df_course 변수와 합쳐 교양 과목 변수(df_course)에 수강 횟수 정보 추가
df_elective = pd.merge(df_elective, df_course, how='inner', on=["전공분야코드", "일련번호", "교과목명"])
# 같은 교과목 코드에 최신 교과목 이외 나머지 교과목을 삭제(교양 학점 계산 목적)
df_elective = df_elective.drop_duplicates(["전공분야코드", "일련번호"], keep='last')
# 컬럼명 재배열(최초개설년도, 최초개설학기 삭제)
df_elective = df_elective[["전공분야코드", "일련번호", "교과목명", "학점", "수강횟수", "분류"]]

# 4-1-2. 예체능 데이터프레임 정의
df_physical_art = df_course[df_course["전공분야코드"] == "GS"]
df_physical_art = df_physical_art[df_physical_art["일련번호"].str.startswith("01") |
                                  df_physical_art["일련번호"].str.startswith("02")]
# 컬럼명 재배열(최초개설년도, 최초개설학기 삭제)
df_physical_art = df_physical_art[["전공분야코드", "일련번호", "교과목명", "학점", "수강횟수"]]
# 같은 교과목 코드에 최신 교과목 이외 나머지 교과목을 삭제(교양 학점 계산 목적)
df_physical_art = df_physical_art.drop_duplicates(["전공분야코드", "일련번호"], keep='last')
# 예체능 과목의 "분류" 컬럼명 새로 추가
df_physical_art["분류"] = "pna"

# 4-1-3. 연구 과목 데이터프레임 정의
list_tf_res1 = df_course["일련번호"] == "9102"
list_tf_res2 = df_course["일련번호"] == "9103"
df_research = df_course[list_tf_res1 | list_tf_res2]
# 컬럼명 재배열(최초개설년도, 최초개설학기 삭제)
df_research = df_research[["전공분야코드", "일련번호", "교과목명", "학점", "수강횟수"]]
# 같은 교과목 코드에 최신 교과목 이외 나머지 교과목을 삭제(교양 학점 계산 목적)
df_research = df_research.drop_duplicates(["전공분야코드", "일련번호"], keep='last')
# 예체능 과목의 "분류" 컬럼명 새로 추가
df_research["분류"] = "res"

# 4-1-4. 교양 과목과 예체능 데이터프레임 합산
df_elect_pna_res = pd.concat([df_elective, df_physical_art, df_research])
# 교양-예체능-연구 분류코드 리스트 생성
list_elect_pna_res_code = list(dict.fromkeys(df_elect_pna_res["분류"].values.tolist()))


# 4-2. 대학 전공별 데이터베이스
# 대학 전공분야코드에 따라 선별한 개설강좌정보
df_undergraduate = df_course[df_course["전공분야코드"].isin(list_undergraduate_major_code)]
# 다음 조건에 따라 정렬
df_undergraduate = df_undergraduate.sort_values(by=["최초개설년도", "최초개설학기"])
# 같은 교과목명에 최신 교과목 이외 나머지 교과목을 삭제(전공 학점 계산 목적)
df_undergraduate = df_undergraduate.drop_duplicates(["교과목명"], keep='last')
# GS 또는 UC 제외 대학원 및 연구과목(5XXX 이상) 제거
df_undergraduate = df_undergraduate[(df_undergraduate["전공분야코드"] == "GS") |
                                    (df_undergraduate["전공분야코드"] == "UC") |
                                    (df_undergraduate["일련번호"].str[0] <= "4")]
# 다음 조건에 따라 정렬
df_undergraduate = df_undergraduate.sort_values(by='일련번호')
# 컬럼명 재배열(최초개설년도 및 학기 삭제)
df_undergraduate = df_undergraduate[["전공분야코드", "일련번호", "교과목명", "학점", "수강횟수"]]

print("1. 학번(student_number)")
print(student_number)
print("2. 학생 성적(df_student)")
print(df_student)
print("3. 전체 과목 정보")
print(df_course)
print("3-1. 전공분야코드(ex. GS, MM 등)별 전체 과목 설명 정보")
print(df_major_explain)
print("4. 교양-예체능-연구 과목 정보")
print(df_elect_pna_res)
print("4-1. 분류별 교양-예능-연구 과목 정보")
print(df_elect_pna_res_explain)
print("5. 전공분야코드별 학부생 전용 과목 정보")
print(df_undergraduate)
