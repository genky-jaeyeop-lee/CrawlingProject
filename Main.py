# Main.py
from Crawling import crawl_course_details  # Crawling.pyからクローリング関数インポート
from InputExcel import save_to_excel  # InputExcel.pyからExcelに保存する関数インポート
import openpyxl


def get_urls_from_excel(file_path, target_name="李"):
    """ExcelファイルからURLリストを読み込んで返す関数(A列で名前をフィルタリングした後、URLを抽出)"""
    urls = []
    try:
        # Excel 파일 열기
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active

        # A列で担当者名をチェックし、担当者へのURLのみを抽出します
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=9):  # A列からIまで
            name_cell = row[0]  # A列担当者名前
            url_cell = row[8]   # I列 URL
            if name_cell.value == target_name and url_cell.hyperlink:
                urls.append(url_cell.hyperlink.target)  # ハイパーリンク URL 抽出
            elif name_cell.value != target_name:
                continue

        print("추출된 URL 목록:", urls)  # URL リストを出力して上手く抽出されるか確認
        return urls
    except Exception as e:
        print(f"Excel 파일 읽기 중 에러 발생: {e}")
        return []





def main():
    # Input Excelファイルのパス
    excel_file_path = r"C:\Users\L1003613honbu048\Documents\業務\メモ\ピクチャー\QC活動\セミナー研修検討リスト.xlsx"  # 여기에 URL 목록이 있는 엑셀 파일 경로를 넣어주세요

    # ExcelからURLリストを読み込む
    url_list = get_urls_from_excel(excel_file_path)

    if not url_list:
        print("URL リストを修得出来ません.")
        return

    # URL リストを出力して上手く抽出されるか確認
    print("最終 URL リスト:", url_list)

    # 各URLに対してクローリングおよびエクセルファイルを保存します
    for url in url_list:
        print(f"クローリング中: {url}")
        crawled_data = crawl_course_details(url)
        if crawled_data:
            file_path = r"C:\Users\L1003613honbu048\Documents\業務\メモ\ピクチャー\QC活動\course_details.xlsx" #Output Excelファイルのパス
            save_to_excel([crawled_data], file_path)



if __name__ == "__main__":
    main()