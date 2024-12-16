import gspread
from google.oauth2.service_account import Credentials


def update_spreadsheet(data_list, sheet_id, sheet_name):
    """
    스프레드시트에 데이터를 업데이트합니다.

    Args:
        data_list (list): 크롤링한 데이터 리스트 (각 항목이 dict 형태)
        sheet_id (str): 스프레드시트 ID
        sheet_name (str): 업데이트할 시트 이름
    """
    try:
        # Google Sheets 인증
        credentials = Credentials.from_service_account_file('path_to_your_service_account.json')
        gc = gspread.authorize(credentials)

        # 스프레드시트 열기
        sh = gc.open_by_key(sheet_id)
        worksheet = sh.worksheet(sheet_name)

        # 데이터 업데이트
        for idx, data in enumerate(data_list, start=2):  # 헤더 다음부터 입력
            worksheet.update(f'A{idx}', data["詳細概要"])
            worksheet.update(f'B{idx}', "레벨 미정")
            worksheet.update(f'C{idx}', data["受講形式"])

        print("Spreadsheet updated successfully!")
    except Exception as e:
        print(f"Error while updating spreadsheet: {e}")
