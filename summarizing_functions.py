"""
    ZEUS 상에서의 대학 및 대학원 개설강좌정보 통합 및 교과목 정리
"""

import pandas as pd  # 데이터프레임 생성용 패키지


def summarize_course():
    # 1. 개설강좌정보 통합
    course_undergraduate = pd.read_excel("./개설강좌정보_대학.xls")
    course_graduate = pd.read_excel("./개설강좌정보_대학원.xls")
    course = pd.concat([course_undergraduate, course_graduate])

    # 2. 개설강좌정보 중 필요한 컬럼만 추출
    course = course[['년도', '교과목-분반', '교과목명', '강/실/학']]

    course['전공분야코드'] = course['교과목-분반'].str[0:2]
    course['난이도'] = course['교과목-분반'].str[2]
    course['일련번호'] = course['교과목-분반'].str[3:6]
    course['학점'] = course['강/실/학'].str[-1]

    # 3. 필요한 컬럼만 수집 및 중복되는 데이터 제거
    course = course[['년도', '전공분야코드', '난이도', '일련번호', '학점', '교과목명']].drop_duplicates()

    # 4. {최초개설년도 - 교과목 정보} 형태로 묶어서 요약
    course = course.groupby(['전공분야코드', '난이도', '일련번호', '학점', '교과목명'], as_index=False)['년도'].min()
    course = course.rename(columns={'년도': '최초개설년도'})
    course = course[['최초개설년도', '전공분야코드', '난이도', '일련번호', '학점', '교과목명']]

    # print(course.sort_values(by=['최초개설년도']))
    # Final. 요약된 파일을 return
    return course


pd.set_option('display.max_columns', 4)
# def summarize_student_information():
# 1. 성적표 데이터 요약
previous = pd.read_excel("./성적표.xls")

# 1-1. 학번 추출
student_number = previous.iloc[0][0].strip()[-8:]

# 1-2. 의미없는 열, 행 및 결측치(NaN) 제거
previous = previous.drop([previous.columns[0], previous.columns[2]], axis=1)  # 열 제거
previous = previous.dropna(axis=0, how='any')  # 결측치 제거
previous = previous.drop(previous.index[0], axis=0)  # 맨 처음 행(header data) 제거

# 1-3. 인덱스 리셋 및 컬럼명 변경
previous = previous.reset_index(drop=True)
previous.columns = ['과목코드', '과목명', '학점', '평점']

# 1-4.
print(previous)
# 2. 현재수강과목 데이터 요약
present = pd.read_excel("./현재수강과목.xls")

