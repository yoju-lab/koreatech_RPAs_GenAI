# 네이버 검색 API 예제 - 블로그 검색
import os
import sys
import urllib.request

from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

# .env에서 값 가져오기
client_id = os.getenv("NAVER_CLIENT_ID")
client_secret = os.getenv("NAVER_CLIENT_SECRET")

encText = urllib.parse.quote("포켄스")
# url = "https://openapi.naver.com/v1/search/blog.xml?query=" + encText # XML 결과
# url = "https://openapi.naver.com/v1/search/shop?sort=date&query=" + encText # JSON 결과
# url = "https://openapi.naver.com/v1/search/shop?display=20&sort=date&query=" + encText # JSON 결과
# url = "https://openapi.naver.com/v1/search/news?display=20&sort=date&query=" + encText # JSON 결과
url = "https://openapi.naver.com/v1/search/image?display=20&sort=date&query=" + encText # JSON 결과
request = urllib.request.Request(url)
request.add_header("X-Naver-Client-Id",client_id)
request.add_header("X-Naver-Client-Secret",client_secret)
response = urllib.request.urlopen(request)
rescode = response.getcode()
if(rescode==200):
    response_body = response.read()
    print(response_body.decode('utf-8'))
else:
    print("Error Code:" + rescode)