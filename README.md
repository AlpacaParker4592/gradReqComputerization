# Graduation Requirements Computerization Project *for Undergraduates*
<div align="center">
<img src="https://img.shields.io/badge/Python-3766AB?style=flat&logo=Python&logoColor=white"/>
<img src="https://img.shields.io/badge/MS Excel 2016 or later-217346?style=flat&logo=Microsoft Excel&logoColor=white"/>
</div><br>

<div align="center">

![result](https://user-images.githubusercontent.com/63055303/140486999-ba639f3d-0bb3-4f75-82a5-94164020c6bb.png)

</div>

이전의 [졸업 요건 전산화 프로그램](https://github.com/AlpacaParker4592/GIST_Credit_Analysis_Program_without_IDE) 
배포와 관련하여 관심을 가져주신 모든 분들께 감사드립니다.
해당 프로그램은 이전과 같이 성적표 파일과 현재 수강 중인 과목 파일을 바탕으로 졸업요건을 전산화하여
현재 졸업요건 충족 여부를 확인하는 프로그램입니다.

기존 프로그램에는 웹 사이트 형태로 변형하여 제공한다고 했으나,
개인적인 사정으로 기간 내에 프로젝트를 완료할 수 없다고 판단하여
대신 프로그램 형태를 콘솔 형태로 진행된다는 점 양해 부탁드립니다.

## Information Provided
+ 학번(15학번 이상) 및 전공별 졸업요건 만족 현황 및 학점 이수현황
+ 각 교과목 정보 및 수강횟수(코드쉐어 적용 과목 포함)

## Execution Process
1. Code 버튼을 눌러 해당 폴더를 zip 파일로 내려받은 후 압축을 해제하십시오.
2. 프로그램 실행 전 다음 파일을 괄호 안의 이름으로 Data 폴더 내에 저장해야 합니다.
    + ZEUS 내 xls 파일 다운로드 방법은 하단의 ***[Precaution for Users](#precaution-for-users)*** 문단을 참고 바랍니다.

    > + **기본 제공 파일**
    >    1. 템플릿 파일(template.xlsm(자동 필터링 기능 있음), template.xlsx(자동 필터링 기능 없음))
    >    2. 교양 과목 목록 파일(elective_course_list.xlsx)
    > 
    > + **필수 저장 파일**
    >    1. 2015학년 1학기부터의 대학 과목 정보 파일(course_information_undergraduate.xls)
    >        - ZEUS의 '**수업 > 개설강좌조회**' 탭 선택 후 대학분류를 '**GIST대학**'으로 조회하여 파일 다운로드
    >    2. 2015학년 1학기부터의 대학원 과목 정보 파일(course_information_graduate.xls)
    >        - ZEUS의 '**수업 > 개설강좌조회**' 탭 선택 후 대학분류를 '**대학원**'으로 조회하여 파일 다운로드
    >    3. 성적 관련 파일(grade_report.xls)
    >        - ZEUS의 '**성적 > 개인성적조회**' 탭 선택 후 '**Report Card**' 버튼으로 성적표 다운로드(Kor, Eng 선택 무관)
    >
    > + **선택 저장 파일**
    >    1. 현재 수강 과목 정보 파일(present_course_registration.xls)
    >        - ZEUS의 '**수업 > 수강신청내역조회**' 탭 선택 후 파일 다운로드

3. 압축 해제한 폴더 내에서 exe 파일 실행 후 생성된 파일(*computerization_result_kor.xlsm 또는 .xlsx*)을 확인하시면 만족 현황을 조회하실 수 있습니다.

## Precaution for Users
### 프로그램 이용 관련
#### ZEUS 내 xls 파일 다운로드
<div align="center">

![save_xls_file](https://user-images.githubusercontent.com/63055303/140265210-bd61aba6-e79f-4e3f-b37b-89ad84fdd88a.png)

</div>

+ 과목 정보 파일 등 엑셀 파일 저장 시 위 그림과 같이 테이블 상에서 우클릭 후 '엑셀 저장' 버튼을 눌러 저장할 수 있습니다.
그 결과 data 폴더 내 파일 구성이 다음과 같이 이루어져야 합니다.

<div align="center">

![file_configuration](https://user-images.githubusercontent.com/63055303/140388776-6dacf095-40f3-4c4b-a417-626969110e62.png)

</div>

#### 그 이외
+ **해당 프로그램에서 나온 결과는 귀하의 졸업을 보장하지 않습니다.** 참고용으로만 사용하시길 바랍니다.
+ 현재 테스트가 충분히 이뤄지지 않아 버그가 있을 수 있습니다. 이 점 양해 부탁드리며 버그 발생 시 아래 이메일로 제보 주시면 감사하겠습니다.
+ 현재 **15학번 ~ 21학번 학부생**을 대상으로 사용 가능합니다. 교양 과목 분류 문제로 14학번 이전 학부생은 사용이 원활하지 않을 수 있습니다.
+ 이 프로그램은 귀하의 졸업 요건을 그래프로 나타내기 쉽도록 졸업 요건 일부를 통합하였습니다. 
자세한 졸업 요건이나 지금까지 이수한 교과목 전체를 알고자 하는 경우에는 다음 URL을 참고하시길 바랍니다.
  + [광주과학기술원 학사편람(2021년판)](https://college.gist.ac.kr/college/sub03_01_05_10_10.do)

### 수강 과목 관련
+ 교양 과목 분류(HUS, PPE, GSC)의 경우 각 연도별 학사편람을 참고하여 제작하였습니다.
이때 연도에 따라 과목 분류가 달라지는 경우가 발견됐습니다 (ex. 강대국의 흥망: 2018년에 GSC 과목에서 HUS 과목으로 변경). 이에 대해서는 다음과 같습니다.
    > 분류가 잘못된 과목이 있을 수 있습니다.
      예를 들어 'GS2601: 동아시아의 전통과 현대' 과목의 경우 2016 학사편람 기준으로는 GSC로 분류,
      2017 - 2020 학사편람 기준으로는 HUS로 분류됩니다...(후략)
    >+ (2020.07.10) 학사지원팀에 이에 대해 문의해본 결과 그러한 과목 모두에 일률적으로 적용되는 규칙은 없다는 말씀을 전달받았습니다
       (각각 해당 수업을 수강한 년도의 학사편람을 따를수도 있으며 분류 중 한 가지를 자유롭게 선택할 수도 있는 등)...(후략)
    >
    > [Google Colab 졸업요건 분석 프로그램](https://colab.research.google.com/drive/1pRaZLyTsbN9RIpmoCs-645dxTWQDM_LQ?usp=sharing&fbclid=IwAR0yx6ptBulpYTaRz9zea9JW7H617tWE518gcrUqDlzWDYFdH73gwfopQ-A)
      설명에서 일부 인용

+ 해당 프로그램은 일부 교양 과목이 GSC 과목에서 HUS 또는 PPE 과목으로 전환된 경우만 있는 점을 들어 이전에 GSC 과목으로 이수했더라도
이후 HUS 또는 PPE 과목으로 전환됐을 경우 전환된 코드로 변경되도록 하였습니다.
    + ex) "강대국의 흥망" 과목을 2017년 이전에 이수했어도 2018년 이후에 HUS 과목으로 변경됐으므로 **HUS 과목을 이수한 것으로 처리**
+ 다음 알고리즘에 따라 대부분의 교과목에 대해 코드쉐어를 반영하고 있습니다.
빠진 과목이 있는 경우 메일을 통해 제보해주시면 최대한 빠른 시일 내에 수정하겠습니다.
    > **코드쉐어 과목 추출 알고리즘**
    > 1. 성적표 및 현재 수강 과목 정보 DB에서 추출한 과목 중 하나와 **교과목 코드(알파벳+4자리 일련번호)가 일치하는 과목들**을 과목 정보 DB에서 추출함.
    >     + '유체역학' 과목을 **MC2102**로 수강(이수)했을 경우 과목 정보 DB에서 해당 코드와 같은 과목을 추출
            (21년도 2학기 기준 '**MC2102**: 유체역학', '**MC2102**: 유체역학 I' 두 과목 존재)
    > 2. 1에서 추출한 과목과 **과목명이 같은 과목들**을 다시 과목 정보 파일에서 추출함.
    >     + 'MC2102: **유체역학**' - 'MC3105: **유체역학**', 'EV2210: **유체역학**', 'EV3218: **유체역학**'
    >     + 'MC2102: **유체역학 I**' - 'EV2210: **유체역학 I**'
    > 3. 2에서 추출한 과목들(1에서 추출한 과목 포함)의 이수 횟수에 1을 더함.
    > 4. 수강 또는 이수한 과목을 모두 확인할 때까지 과정 1부터 3까지를 반복함.
+ 현재 전공 제한 학점(17학번 이전: 36학점, 18학번 이후: 42학점)을 초과하여 전공 과목을 수강한 경우 **타 전공 코드쉐어 과목의 적용과 무관하게 모두 절삭하여 반영**하고 있습니다.
이 점 양해 부탁드립니다. 자세한 내용은 [**여기**](https://github.com/AlpacaParker4592/gradReqComputerization/issues/1)에서 확인하실 수 있습니다.
+ 해당 프로그램은 허락 없이 포크하여 사용하셔도 됩니다. 
사용한 변수는 [Wiki 탭](https://github.com/AlpacaParker4592/gradReqComputerization/wiki)의
"**더 나은 프로그램을 제작하고자 하시는 분들을 위한 도움말**(파일 및 변수 설명)" 항목에서 확인하실 수 있으며,
제작하실 때는 Issues 탭의 내용을 고려하시면 되겠습니다.
+ 실행 중 오류가 발생하거나 나온 결과가 실제 반영되는 학점과 다른 사항을 발견할 시
[**Github Issues**](https://github.com/AlpacaParker4592/gradReqComputerization/issues) 탭을 이용하거나(권장)
**lhh-znso4(at)gm.gist.ac.kr**로 연락해주시길 바랍니다.
