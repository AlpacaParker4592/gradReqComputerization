"""
프로그램 관련 함수 모음집
"""
import pandas as pd  # 데이터프레임 생성용 패키지
import os


def tf_exist_all_files():
    """
    파일 존재 여부 확인 함수
    :return:
    0(모든 파일 존재) 또는 1(대학원생 교과목 정보 파일만 존재) 또는
    2(현재 수강 교과목 정보 파일만 존재) 또는 3(필수 파일만 존재) 또는
    4(필수 파일이 존재하지 않음)
    """
    file_path = "./"
    file_list = os.listdir(file_path)

    # 학부생 교과목 정보 파일(course_information_undergraduate.xls): 필수
    # 대학원생 교과목 정보 파일(course_information_graduate.xls): 필수
    # 교양 과목 정보 파일(elective_course_list.xlsx): 필수
    # 성적 정보 파일(grade_report.xls): 필수
    # 현재 수강 교과목 정보 파일(present_course_registration.xls): 선택

    tf_exist_file_e1 = "course_information_undergraduate.xls" in file_list  # 학부생 교과목 정보 파일 존재 여부
    tf_exist_file_e2 = "course_information_graduate.xls" in file_list  # 대학원생 교과목 정보 파일 존재 여부
    tf_exist_file_e3 = "elective_course_list.xlsx" in file_list  # 교양 과목 정보 파일 존재 여부
    tf_exist_file_e4 = "grade_report.xls" in file_list  # 성적 정보 파일 존재 여부
    tf_exist_file_e5 = "template.xlsx" in file_list  # 템플릿 파일 존재 여부
    tf_exist_file_s1 = "present_course_registration.xls" in file_list  # 현재 수강 교과목 정보 파일

    # 필수 파일 존재 여부 확인
    print("학부생 교과목 정보 파일:", end="\t")
    if tf_exist_file_e1:
        print("YES")
    else:
        print("NO")
        print("학부생 교과목 정보 파일(course_information_undergraduate.xls)이 존재하지 않습니다.")
        return 2

    print("대학원생 교과목 정보 파일:", end="\t")
    if tf_exist_file_e2:
        print("YES")
    else:
        print("NO")
        print("대학원생 교과목 정보 파일(course_information_graduate.xls)이 존재하지 않습니다.")
        return 2

    print("교양 과목 정보 파일:", end="\t\t")
    if tf_exist_file_e3:
        print("YES")
    else:
        print("NO")
        print("교양 과목 정보 파일(elective_course_list.xlsx)이 존재하지 않습니다.")
        return 2

    print("성적 정보 파일:", end="\t\t\t")
    if tf_exist_file_e4:
        print("YES")
    else:
        print("NO")
        print("성적 정보 파일(grade_report.xls)이 존재하지 않습니다.")
        return 2

    print("템플릿 파일:", end="\t\t\t")
    if tf_exist_file_e5:
        print("YES")
    else:
        print("NO")
        print("템플릿 파일(template.xls)이 존재하지 않습니다.")
        return 2

    # 선택 파일 존재 여부 확인
    print("현재 수강 교과목 정보 파일:", end="\t")
    if tf_exist_file_s1:
        print("YES")
    else:
        print("NO")
        return 1

    return 0


def summarize_course(ex_num):
    """
    ZEUS 상에서의 대학 및 대학원 개설강좌정보 통합 및 교과목 정리 함수
    :return:
    2016년부터 현재까지 개설된 과목 정보 dataframe
    """
    # 1. 개설강좌정보 통합
    tf_exist_graduate = ex_num == 0 or ex_num == 1  # 대학원 개설강좌정보 파일 존재 유무

    course_undergraduate = pd.read_excel("./course_information_undergraduate.xls")
    course = course_undergraduate
    # 대학원 개설강좌정보 파일이 존재할 경우에만 추가하여 합치기
    if tf_exist_graduate:
        course_graduate = pd.read_excel("./course_information_graduate.xls")
        course = pd.concat([course_undergraduate, course_graduate])

    # 2. 개설강좌정보 중 필요한 컬럼만 추출 및 컬럼 내 데이터 추출
    course = course[['년도', '학기', '교과목-분반', '교과목명', '강/실/학']].copy()

    course['전공분야코드'] = course['교과목-분반'].str[0:2]
    course['난이도'] = course['교과목-분반'].str[2]
    course['일련번호'] = course['교과목-분반'].str[3:6]
    course['학점'] = course['강/실/학'].str[-1].astype(int)

    # 3. 필요한 컬럼만 수집 및 중복되는 데이터 제거
    course = course[['년도', '학기', '전공분야코드', '난이도', '일련번호', '학점', '교과목명']].drop_duplicates()

    # 3+. 학기 명칭 변경
    course.loc[course['학기'] == '1학기', '학기'] = '1'
    course.loc[course['학기'] == '여름학기', '학기'] = '2'
    course.loc[course['학기'] == '2학기', '학기'] = '3'
    course.loc[course['학기'] == '겨울학기', '학기'] = '4'
    course.loc[course['학기'].str.len() != 1, '학기'] = '5'  # 4개 학기에 포함되지 않는 학기(ex. 인정학기 등)는 5로 처리

    # 4. {최초개설년도 - 교과목 정보} 형태로 묶어서 요약
    course = course.groupby(['전공분야코드', '난이도', '일련번호', '학점', '교과목명'], as_index=False)[['년도', '학기']].min()
    course = course.rename(columns={'년도': '최초개설년도', '학기': '최초개설학기'})
    course = course[['최초개설년도', '최초개설학기', '전공분야코드', '난이도', '일련번호', '교과목명', '학점']]

    # print(course.sort_values(by=['최초개설년도', '최초개설학기']))
    return course


def summarize_student_information(ex_num):
    """
    성적표 및 현재수강정보에서 학번 및 수강한 과목 정리
    :return:
    1. 학번 8자리
    2. 현재까지 수강 또는 이수한 과목 관련 dataframe
    """
    pd.set_option('display.max_columns', 100)  # 테이블 조회 시 테스트용 명령

    # 1. 성적표 데이터 요약
    previous = pd.read_excel("./grade_report.xls")
    # 1-1. 학번 추출
    student_number = int(previous.iloc[0][0].strip()[-8:])

    # 1-2. 의미없는 열, 행 제거
    previous = previous.drop([previous.columns[0], previous.columns[2]], axis=1)  # 열 제거
    previous = previous.drop(previous.index[:3], axis=0)  # 불필요한 부분 제거 1
    previous = previous.drop(previous.index[-3:], axis=0)  # 불필요한 부분 제거 2

    # 1-3. 인덱스 리셋 및 컬럼명 변경
    previous = previous.reset_index(drop=True)
    previous.columns = ['과목코드', '과목명', '학점', '평점']

    # 1-4. 수강학기 열 추가
    previous['수강연도'] = ""
    previous['수강학기'] = ""
    # 초기 수강 연도 및 학기 설정(AP 수강)
    year = "AP"
    semester = "AP"
    # 각 교과목의 수강 연도 및 학기 추가
    for row in range(len(previous)):
        # 연도 또는 학기가 바뀔 시 갱신
        if previous['과목명'].iloc[row].strip()[0] == "<" and previous['과목명'].iloc[row].strip()[-1] == ">":
            year_semester = previous['과목명'].iloc[row].strip()[1:-1].split("/")
            year = int(year_semester[0])  # 출력값: 2019, 2020, 2021 등
            semester = year_semester[1]  # 출력값: (한글)1학기, 여름학기,... / (영문)Spring Semester,...
        previous['수강연도'].iloc[row] = year
        previous['수강학기'].iloc[row] = semester

    # 1-5. 필요없는 행 제거
    previous = previous.dropna(subset=['과목코드'])
    previous = previous.reset_index(drop=True)

    # 1-6. 컬럼 내 데이터 추출 및 정리
    previous['전공분야코드'] = previous['과목코드'].str[0:2]
    previous['난이도'] = previous['과목코드'].str[2]
    previous['일련번호'] = previous['과목코드'].str[3:6]
    previous['학점'] = previous['학점'].astype(int)
    previous = previous[['수강연도', '수강학기', '전공분야코드', '난이도', '일련번호', '과목명', '학점', '평점']]

    # 1-7. 정리한 성적표 데이터를 수강 기등록 과목 데이터에 넣기
    course_registration = previous

    tf_exist_present = ex_num == 0  # 현재수강과목 파일 존재 유무
    if tf_exist_present:
        # 2. 현재수강과목 데이터 요약
        present = pd.read_excel("./present_course_registration.xls")
        present = present.drop(present.index[0], axis=0)  # 맨 처음 행(header data) 제거

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
            present['학점'] = present["강/실/학"].str[-1].astype(int)
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

    return student_number, course_registration


def summarize_elective_course():
    """
    elective_course_list.xlsx 내 파일 데이터 요약 함수
    :return: 
    교양 과목 관련 요약 dataframe
    """
    # 1. 분류별 데이터 추출
    hus_list = pd.read_excel("./elective_course_list.xlsx", sheet_name="hus")
    hus_list = hus_list[['교과목', '학점', '교과목명']].drop_duplicates()
    hus_list = hus_list.reset_index(drop=True)
    hus_list['분류'] = 'hus'

    ppe_list = pd.read_excel("./elective_course_list.xlsx", sheet_name="ppe")
    ppe_list = ppe_list[['교과목', '학점', '교과목명']].drop_duplicates()
    ppe_list = ppe_list.reset_index(drop=True)
    ppe_list['분류'] = 'ppe'

    gsc_list = pd.read_excel("./elective_course_list.xlsx", sheet_name="gsc")
    gsc_list = gsc_list[['교과목', '학점', '교과목명']].drop_duplicates()
    gsc_list = gsc_list.reset_index(drop=True)
    gsc_list['분류'] = 'gsc'

    # 2. gsc 관련 데이터에서 다른 분류(hus, ppe)와 중복되는 과목 삭제
    # print(len(gsc_list))
    is_gsc_hus_duplicated = pd.concat([gsc_list, hus_list]).duplicated(['교과목', '학점', '교과목명'],
                                                                       keep='last').iloc[0:len(gsc_list)]
    gsc_list = gsc_list[~is_gsc_hus_duplicated]
    # print(len(gsc_list))
    is_gsc_ppe_duplicated = pd.concat([gsc_list, ppe_list]).duplicated(['교과목', '학점', '교과목명'],
                                                                       keep='last').iloc[0:len(gsc_list)]
    gsc_list = gsc_list[~is_gsc_ppe_duplicated]
    # print(len(gsc_list))

    # 3. 모든 분류 통합
    elective_list = pd.concat([hus_list, ppe_list, gsc_list], ignore_index=True)
    elective_list['전공분야코드'] = elective_list['교과목'].str[0:2]
    elective_list['난이도'] = elective_list['교과목'].str[2]
    elective_list['일련번호'] = elective_list['교과목'].str[3:6]
    elective_list = elective_list[['전공분야코드', '난이도', '일련번호', '교과목명', '분류']]  # 학점은 학사편람 기준이므로 제외

    # print(elective_list)
    return elective_list
