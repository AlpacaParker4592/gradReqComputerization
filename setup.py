import pandas as pd  # 데이터프레임 생성용 패키지
import functions as func  # functions.py 파일의 함수 사용

# 1. 필수 파일 존재 여부 확인
existence_number = func.tf_exist_all_files()
if existence_number == 4:
    quit(0)

# 2. 관련 변수 설정
table_course = func.summarize_course(existence_number)
student_number, table_student_course = func.summarize_student_information(existence_number)

