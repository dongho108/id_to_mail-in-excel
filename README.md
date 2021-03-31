
# 프로젝트 시작 동기

작년부터 스마트스토어에서 영어교재를 판매하고있다.

[스마트스토어]: https://smartstore.naver.com/double00k

교재를 구매한 학생들에게 변형문제와 유의어등 시험에 도움이 될만한 자료를 추가로 보내주기로 약속했다.

자료는 메일로 보내줄 것이기 때문에 메일 주소가 필요하다.

매일매일 엑셀로 배송명단을 편집하고있는데 여기서 아이디만 추출해 메일로 변환하려 한다.


![excel](https://user-images.githubusercontent.com/54317630/113139365-fbdd4f00-9261-11eb-8f6c-771861c0663f.png)

현재 엑셀은 이렇게 되어있다.

네이버에서 배송명단을 엑셀로 다운받을 때 아이디만 볼 수 있다.

여기서 @naver.com 만 붙이면 바로 메일로 변환된다.

그래서 엑셀에서 아이디를 추출해 메일로 변경하는 작업이 필요하다.





## 설계
* 엑셀을 읽어와야한다.
* 엑셀에서 특정 열을 읽어야 한다.
* 읽은 열에서 아이디 데이터를 추출해야한다.
* 아이디를 메일로 변환해야한다.
* 날짜별로 다른 엑셀들을 알아서 읽어와야한다.
* 메일로 변환된 데이터들을 100개 단위로 출력해야한다.



## 엑셀 읽기


### 라이브러리
pandas
openpyxl
xlrd


### 변수
prefix : 배송엑셀들을 모아놓은 폴더 접근
date_list : 읽을 엑셀이 필요한 날짜
surfix : .xlsx


### 구현
pandas의 read_excel(경로, sheet_name='시트이름',  engine='openpyxl')를 사용했다.




## 엑셀에서 id 추출하기


### 구현
xlsx.loc[:, ['구매자ID']] : 왼쪽부분은 모든 행을 읽는다는 뜻, 오른쪽 부분은 읽을 열을 정함

xlsx를 출력하면 아래와 같이 나온다.

![pandas_dataframe](https://user-images.githubusercontent.com/54317630/113139377-fe3fa900-9261-11eb-990c-f35ff2effe29.png)

아직 xlsx에는 pandas의 dateframe 이라는 타입으로 되어있기 때문이다.
그래서 여기서 "데이터" 만을 추출해야한다.

ids = xl_ids.values
뒤에 .values를 붙이면 해당 데이터가 2차원 리스트로 반환된다.

![pandas_list](https://user-images.githubusercontent.com/54317630/113139380-fed83f80-9261-11eb-8662-0bb93ec1fca7.png)





## id를 mail로 변환

### 구현
ids에서 @naver.com 을 붙이고 리스트에 담으면된다.



## main함수

### 구현
main에서는 여러 날짜(여러 엑셀파일)을 다 돌려서 메일주소를 받아야한다.
그리고 중복된 메일이 있을 수 있다. 그래서 set에 넣었다가 다시 list로 가져왔다.
네이버 메일은 한번에 최대 100명까지 메일을 보낼 수 있다. 그래서 100개 단위로 나눠서 출력해야한다.



## 출력 결과
![output](https://user-images.githubusercontent.com/54317630/113139385-fed83f80-9261-11eb-9c08-1d3dc8952896.png)

## 어려웠던 부분

### 1.
engine='openpyxl'을 붙이지 않았을 때
"xlrd.biffh.XLRDError: Excel xlsx file; not supported" 오류가 발생했다.
pandas에서 엑셀파일을 어떤 엔진으로 읽을지 정해줘야 했다.



### 2.
pandas의 dataframe 형식에서 "데이터"만을 추출해야했다.
.value 를 붙였는데 
" AttributeError: 'DataFrame' object has no attribute 'value' "
오류가 발생했다.

근데 value가 아니라 values를 붙여야 했다.



## 더 발전할 부분


### SMTP 프로토콜 사용
원래는 smtp 프로토콜을 이용하여 메일 전송까지 한번에 처리하려고 했다.
그런데 네이버의 smtp 프로토콜에 인증이 잘 안되었다.
구글의 smtp로 보내면 성공하지만, 구매자들과 소통하는 계정이 네이버였기 때문에 일단은 보류했다.
어차피 아이디만 추출해도 복사 붙여넣기만 하면 끝이기 때문에 시간이 엄청 줄었다.


### 메모리 절약
엑셀 전체를 읽어서 열을 읽는 것이 아닌 애초에 아이디의 열만 가져오면 메모리 절약을 할 수 있다.
하지만 엑셀파일이 몇 안되기 때문에 당장은 수정하지 않을 예정이다.
