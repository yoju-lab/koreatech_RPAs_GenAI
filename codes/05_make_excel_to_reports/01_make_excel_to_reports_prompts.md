```prompt
업무 순서 맞는 보고서 엑셀 만들기 

-현재 폴더에 알맞은 파일명으로 작성 
- 현재 폴더에 작업(os.path.dirname(os.path.abspath(__file__)))

[업무 순서] 
1. 현재 폴더에 genai_rpa.xlsx 없으면 생성
2. genai_rpa_202503291327.xlsx 백업
(202503291327는 현재시각의 year, month, day, hour, minute, second)
3. genai_rpa.xls > prev_list sheet 제거
4. genai_rpa.xls > now_list sheet를 prev_list로 변경
5. genai_rpa.xls >  now_list sheet 생성
6. naver api function 호출 : 쇼핑목록 데이터 가져오기
7. now_list sheet 내용 update
8. genai_rpa.xlsx 파일 저장
```

```python