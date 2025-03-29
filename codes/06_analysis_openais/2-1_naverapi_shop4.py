import xlwings as xw
import datetime
import shutil
import pandas as pd
import json
import os
import sys
import urllib.request
from openai import OpenAI

client_id = ""
client_secret = ""
encText = urllib.parse.quote("포켄스")

# OpenAI API 호출
client = OpenAI(
api_key='',
)
openai_model = "gpt-4o-mini"

# 1. genai_rpa.xlsx 파일 열기
file_path = 'genai_rpa.xlsx'
wb = xw.Book(file_path)


def convert_json_to_dataframe(json_result):
    """
    json_result가 문자열이면 dict로 변환한 후,
    "items" 데이터를 pandas DataFrame으로 변환하고,
    "순위" 열을 제일 왼쪽에 추가하여 1부터 일련번호를 부여한 후
    이 열을 index로 설정합니다.
    """
    if isinstance(json_result, str):
        json_result = json.loads(json_result)
    items = json_result.get('items', [])
    df = pd.DataFrame(items)
    # "순위" 열을 추가 (1부터 시작하는 일련번호)
    df.insert(0, "순위", range(1, len(df)+1))
    # "순위" 열을 index로 설정
    df.set_index("순위", inplace=True)
    return df


def get_naver_shopping_list_data():
    # display를 20으로 수정. 페이징을 한다면 start=번호 형태를 추가
    url = "https://openapi.naver.com/v1/search/shop?sort=date&display=20&query=" + encText # JSON 결과

    request = urllib.request.Request(url)
    request.add_header("X-Naver-Client-Id",client_id)
    request.add_header("X-Naver-Client-Secret",client_secret)
    response = urllib.request.urlopen(request)
    rescode = response.getcode()
    if(rescode==200):
        response_body = response.read()
        # print(response_body.decode('utf-8'))
    else:
        print("Error Code:" + rescode)
    result = response_body.decode('utf-8')

    return result


def handle_list_sheet():
    # 2. 현재 시각을 기반으로 백업 파일 생성
    timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    backup_file = f"genai_rpa_{timestamp}.xlsx"
    shutil.copy(file_path, backup_file)
    print(f"Backup created: {backup_file}")

    # 3. 'prev_list' 시트 삭제 (존재하는 경우)
    sheet_names = [s.name for s in wb.sheets]
    if 'prev_list' in sheet_names:
        wb.sheets['prev_list'].delete()
        print("Sheet 'prev_list' deleted.")
    else:
        print("Sheet 'prev_list' does not exist. Skipping deletion.")

    # 4. 'now_list' 시트의 이름을 'prev_list'로 변경
    if 'now_list' in sheet_names:
        wb.sheets['now_list'].name = 'prev_list'
        print("Sheet 'now_list' renamed to 'prev_list'.")
    else:
        print("Sheet 'now_list' not found. Please check the workbook.")
        wb.close()
        return

    # 5. 새로운 'now_list' 시트 생성
    new_sheet = wb.sheets.add(name="now_list", after=wb.sheets['prev_list'])
    print("Sheet 'now_list' created.")

    return


def update_now_list(df_shopping):
    # 8. 'now_list' 시트 내용 업데이트 (기존 내용 클리어 후 A1셀부터 DataFrame 쓰기)
    now_sheet = wb.sheets['now_list']
    now_sheet.clear()
    now_sheet.range('A1').value = df_shopping
    print("Sheet 'now_list' updated with new shopping data.")

    return


def conn_openai_api(prompt):
    completion = client.chat.completions.create(
        model=openai_model,
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )

    # 분석글 출력
    print(completion.choices[0].message.content)
    result = completion.choices[0].message.content

    return result


def get_openai_shopping_list_anaysis():
    # 각 시트의 전체 데이터를 불러오기 (테이블 형태이므로 used_range로 모든 셀을 가져옴)
    prev_sheet = wb.sheets['prev_list']
    now_sheet = wb.sheets['now_list']

    prev_data = prev_sheet.used_range.value
    now_data = now_sheet.used_range.value

    # 비교 분석을 위한 프롬프트 구성
    prompt = f"""
    다음 두 목록을 비교 분석하여, prev_list 목록에서 now_list 목록으로 바뀐 주요 특징을 도출해줘.
    prev_list 목록: {prev_data}
    now_list 목록: {now_data}

    비교 분석 결과를 바탕으로 구체적인 수치, 상품명, 쇼핑몰명 등을 언급하며 한글로 500자 이내의 분석글을 작성해줘.
    markdown 언어나, html 태그와 html 특수기호 등을 사용하지 말아줘.
    """
    result = conn_openai_api(prompt)

    return result


def update_now_report(analysis):
    now_sheet = wb.sheets['now_report']

    # 현재 날짜와 시간(년-월-일 시:분:초) 포맷팅
    current_dt = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # A3 셀에 값 입력
    now_sheet.range("A3").value = current_dt+" 기준"
    # A3 셀을 오른쪽 정렬 (엑셀의 상수 xlRight = -4152)
    now_sheet.range("A3").api.HorizontalAlignment = -4152

    now_sheet.range('A5').value = analysis
    now_sheet.range('A5').api.WrapText = True  # 줄바꿈 활성화
    print("Sheet 'now_report' updated with new analysis.")

    return


def save_close_file():
    # 9. genai_rpa.xlsx 파일 저장 및 종료
    wb.save(file_path)
    wb.close()
    print("Workbook saved and closed.")

    return



def main():
    #2~5
    handle_list_sheet()

    # 6. 네이버 API 함수를 호출하여 쇼핑 목록 데이터 가져오기
    result_json = get_naver_shopping_list_data()

    # 7. result JSON을 pandas DataFrame 형태로 만들기 (순위 열 추가됨)
    df_shopping = convert_json_to_dataframe(result_json)
    print("Converted JSON to DataFrame.")

    # 8. 'now_list' 시트 내용 업데이트 (기존 내용 클리어 후 A1셀부터 DataFrame 쓰기)
    update_now_list(df_shopping)

    result_analysis = get_openai_shopping_list_anaysis()

    update_now_report(result_analysis)

    # 9. genai_rpa.xlsx 파일 저장 및 종료
    save_close_file()


if __name__ == '__main__':
    main()