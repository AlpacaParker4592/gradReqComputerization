import pandas as pd  # 데이터프레임 생성용 패키지


def summarize_course():
    """
        ZEUS 상에서의 대학 및 대학원 개설강좌정보 통합 및 교과목 정리
    """
    # 1. 개설강좌정보 통합
    course_undergraduate = pd.read_excel("./course_information_undergraduate.xls")
    course_graduate = pd.read_excel("./course_information_graduate.xls")
    course = pd.concat([course_undergraduate, course_graduate])

    # 2. 개설강좌정보 중 필요한 컬럼만 추출 및 컬럼 내 데이터 추출
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


def summarize_student_information():
    """
        성적표 및 현재수강정보에서 학번 및 수강한 과목 정리
    """
    # pd.set_option('display.max_columns', 100)  # 테이블 조회 시 테스트용 명령

    # 1. 성적표 데이터 요약
    previous = pd.read_excel("./grade_report.xls")

    # 1-1. 학번 추출
    student_number = previous.iloc[0][0].strip()[-8:]

    # 1-2. 의미없는 열, 행 및 결측치(NaN) 제거
    previous = previous.drop([previous.columns[0], previous.columns[2]], axis=1)  # 열 제거
    previous = previous.dropna(axis=0, how='any')  # 결측치 제거
    previous = previous.drop(previous.index[0], axis=0)  # 맨 처음 행(header data) 제거

    # 1-3. 인덱스 리셋 및 컬럼명 변경
    previous = previous.reset_index(drop=True)
    previous.columns = ['과목코드', '과목명', '학점', '평점']
    # print(previous)

    # 1-4. 컬럼 내 데이터 추출 및 정리
    previous['전공분야코드'] = previous['과목코드'].str[0:2]
    previous['난이도'] = previous['과목코드'].str[2]
    previous['일련번호'] = previous['과목코드'].str[3:6]
    previous = previous[['전공분야코드', '난이도', '일련번호', '과목명', '학점', '평점']]

    # 2. 현재수강과목 데이터 요약
    present = pd.read_excel("./present_course_registration.xls")
    present = present.drop(previous.index[0], axis=0)  # 맨 처음 행(header data) 제거

    # 2-1. 필요한 열만 추출, 컬럼명 변경 및 필요없는 행 제거
    present_retake = present.iloc[:, [12, 13, 14]]  # 재수강과목의 이전 수강 과목
    present_retake.columns = ['과목코드', '과목명', '드랍여부']
    present_retake = present_retake.dropna(axis=0, how='all')  # 행 전체가 결측치인 행 제거

    present = present.iloc[:, [1, 2, 7, 14]]  # 현재 수강 과목
    present.columns = ['과목코드-분반', '과목명', '강/실/학', '드랍여부']
    present = present.dropna(axis=0, how='all')  # 행 전체가 결측치인 행 제거

    # 2-2. 드랍한 교과목 제거('드랍여부' 항목이 비어있는 행만 추출)
    present = present[present['드랍여부'].isna()]
    present_retake = present_retake[present_retake['드랍여부'].isna()]

    # 2-3. 컬럼 내 데이터 추출 및 정리
    if len(present) > 0:
        present['전공분야코드'] = present['과목코드-분반'].str[0:2]
        present['난이도'] = present['과목코드-분반'].str[2]
        present['일련번호'] = present['과목코드-분반'].str[3:6]
        present['학점'] = present["강/실/학"].str[-1]
        present = present[['전공분야코드', '난이도', '일련번호', '과목명', "학점"]]
    else:
        present['전공분야코드'] = ""
        present['난이도'] = ""
        present['일련번호'] = ""
        present['과목명'] = ""
        present['학점'] = ""
        present = present[['전공분야코드', '난이도', '일련번호', '과목명', "학점"]]

    if len(present_retake) > 0:
        present_retake['전공분야코드'] = present_retake['과목코드'].str[0:2]
        present_retake['난이도'] = present_retake['과목코드'].str[2]
        present_retake['일련번호'] = present_retake['과목코드'].str[3:6]
        present_retake = present_retake[['전공분야코드', '난이도', '일련번호', '과목명']]
    else:
        present_retake['전공분야코드'] = ""
        present_retake['난이도'] = ""
        present_retake['일련번호'] = ""
        present_retake['과목명'] = ""
        present_retake = present_retake[['전공분야코드', '난이도', '일련번호', '과목명']]

    # 3. 성적표 데이터(previous)와 현재 수강 과목(present) 통합
    course_registration = previous
    if len(present) > 0:
        course_registration = pd.concat([previous, present], ignore_index=True)

    # 4. 현재 재수강하는 과목 데이터 제거(과목코드 기반)
    for row_num in range(len(present_retake)):
        is_same_major_code = course_registration['전공분야코드'] == present_retake.iloc[row_num]['전공분야코드']
        is_same_difficulty = course_registration['난이도'] == present_retake.iloc[row_num]['난이도']
        is_same_num_code = course_registration['일련번호'] == present_retake.iloc[row_num]['일련번호']

        # 재수강하는 과목이 여러 번 이수한 과목(예체능, 콜로퀴움 등)일 경우
        # 재수강 과목 중 처음 C0 이하 또는 U로 이수한 과목 하나를 삭제하도록 조정
        tf_list = ~(is_same_major_code & is_same_difficulty & is_same_num_code)
        found_removal_obj = False  # 삭제 대상 발견 tf값(발견 전: false, 발견 후: true)
        for tf_num in range(len(tf_list.tolist())):
            # 제거 대상 과목 발견 이후의 모든 과목을 그대로 보존
            if found_removal_obj:
                tf_list[tf_num] = True
                continue
            # 수강 과목 목론에서 제거 대상 과목 발견 시 found_removal_obj 값을 true 값으로 바꿈
            # type(tf_list[tf_num]) -> numpy.bool(bool 타입이 아니므로 not 대신 ~ 붙임)
            if ~tf_list[tf_num] and course_registration.iloc[tf_num]['평점'] in ['C0', 'D+', 'D0', 'F', 'U']:
                found_removal_obj = True

        # 제거 대상 과목을 발견하지 못했을 시(평점이 재이수 기준 이상인 경우)
        # print 명령어로 알리고 삭제 대상 과목을 복원
        if not found_removal_obj:
            not_found_name = present_retake.iloc[row_num]['과목명']
            not_found_code = present_retake.iloc[row_num]['전공분야코드'] + \
                             present_retake.iloc[row_num]['난이도'] + \
                             present_retake.iloc[row_num]['일련번호']
            print('ALERT:  드롭 과목 중 과목명이 ' + not_found_name + '(과목 코드: ' + not_found_code + ')인 과목을 제거하지 못했습니다.')
            for i in range(len(tf_list.tolist())):
                tf_list[i] = True
        course_registration = course_registration[tf_list]

    print(course_registration)
    # Final. 학번 및 이수+미이수 교과목 요약본을 return
    return student_number, course_registration
