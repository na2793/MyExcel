# MyExcel

MyExcel은 Android 기반의 spreadsheet application입니다.

여기에는 Android Studio에서 작업한 project 전체가 공개되어 있습니다.<br>
공부한 내용을 점검하기 위한 습작으로, 개발에 적용한 개념이 통상의 개념과 다를 수 있습니다.

## 개요

MyExcel의 인터페이스와 액션은 한글과컴퓨터 사의 한셀과 Microsoft 사의 Excel을 참고하여 구현하였습니다.

### 기능
+ WorkBook을 생성하고, Sheet와 Cell을 편집할 수 있습니다.
+ 작업한 내용을 파일로 저장하고, 다시 불러올 수 있습니다.
+ Tab(선택), Zoom in/out 등의 모션 이벤트를 지원합니다.

### 제약 조건
+ View 클래스를 상속받은 하나의 Custom View로 구현합니다.
+ Csv 포맷을 임의 변형한 Hana 포맷(*.hana)을 이용하여 데이터를 입출력합니다.<br>
(데이터 입출력 테스트를 위한 Hana 파일은 [sample 폴더](https://github.com/na2793/MyExcel/tree/master/sample) 안에 있습니다.)
+ SUM, AVG 함수를 지원합니다.
+ 한글 입력은 고려하지 않습니다.

### 미구현 사항
+ 영역에 대한 연산은 지원하지 않습니다.<br>
(ex. SUM(A1, B2), AVG(A1:A5))
+ Cell에 대한 Context Menu를 제공하지 않습니다.<br>
(복사, 붙여넣기, 채우기 등을 지원하지 않습니다.)
+ 입력된 데이터가 가지고 있는 잠재적인 문제, 유효성 등을 검사하지 않습니다.

### 파악된 이슈
+ View를 invalidate 시 인자로 전달한 invalidate 영역의 값이 onDraw 호출 이전에 바뀌어버립니다.<br>
(무조건 전체 영역을 invalidate 하도록 임시 조치되어 있습니다.)
+ Zoom event가 두 터치 포인트 사이의 중심점이 아닌 현재 보이는 화면의 원점을 기준으로 발생합니다.


## 프로젝트 구조

app > src 이하 activity를 포함한 java 파일들의 구조에 대해서만 서술합니다.<br>
실제 경로는 [여기](https://github.com/na2793/MyExcel/tree/master/app/src/main/java/com/study/hancom/myexcel)를 참조하세요.

###activity
> Activity는 화면에 View를 그리고, View에 Controller를 연결하며, Intent를 통해 입력 받은 데이터를 Controller에 다시 전달합니다. 현재 단일 Activity(Main Activity)로 구현되어 있습니다.

###controller
> Controller는 각각 독립적으로 존재하는 View와 Model을 연결하고, Model의 갱신된 데이터를 View가 알 수 있도록 합니다. 

###model
> Model은 데이터들의 관계와 그 데이터들을 처리하는 방식에 대해 정의합니다.

###view
> View는 화면을 어떻게 구성하고 그려낼 지를 정의합니다.

###util
> Util은 보조적인 작업을 수행합니다. custom listener와 exception을 포함하고 있습니다.
** 개념 정립 및 재분류 필요

## 시스템 구조

(아래부터 추가해야 할 내용)<br>
MVC 구조, Listener의 원리, Hana의 구조<br>
각 모델이 데이터를 어떤 방식으로 관리하고 있는가<br>
중복을 피하기 위해 취한 방법<br>
Cell은 immutable 객체
