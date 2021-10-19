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

# 3. 강좌개설정보에 수강한 강좌 반영
"""
코드쉐어 과목 처리 방법
1. 강좌개설정보(df_course)에서 학생 수강 과목(df_student) 중 하나와 과목코드[전공분야코드, 난이도, 일련번호]가 일치하는 과목을 추출함.
2. 강좌개설정보에서 다음 조건을 만족하는 과목을 추가로 추출함.
    2-1. 1에서 추출한 과목과 [교과목명]이 같은 과목(자신 포함).
3. 2에서 만족하는 과목은 수강 횟수에 1을 더함.
"""
# 3-1. 강좌개설정보 변수(table_course)에 수강 횟수 정보 추가
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
