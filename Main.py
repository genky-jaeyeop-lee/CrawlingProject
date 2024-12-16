# Main.py
from Crawling import crawl_course_details  # Crawling.py에서 크롤링 함수 임포트
from InputExcel import save_to_excel  # InputExcel.py에서 엑셀 저장 함수 임포트
import openpyxl


def get_urls_from_excel(file_path, target_name="李"):
    """Excel 파일에서 URL 목록을 읽어 반환하는 함수 (A열에서 이름 필터링 후 URL 추출)"""
    urls = []
    try:
        # Excel 파일 열기
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active

        # A열에서 담당자 이름을 체크하고 해당 담당자에 대한 URL만 추출
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=9):  # A열에서 I열까지
            name_cell = row[0]  # A열 담당자 이름
            url_cell = row[8]   # I열 URL
            if name_cell.value == target_name and url_cell.hyperlink:
                urls.append(url_cell.hyperlink.target)  # 하이퍼링크 URL 추출
            elif name_cell.value != target_name:
                continue  # 이름이 다르면 해당 행은 무시

        print("추출된 URL 목록:", urls)  # URL 목록을 출력하여 확인
        return urls
    except Exception as e:
        print(f"Excel 파일 읽기 중 에러 발생: {e}")
        return []





def main():
    # URL이 저장된 스프레드시트 경로
    excel_file_path = r"C:\Users\L1003613honbu048\Documents\業務\メモ\ピクチャー\QC活動\セミナー研修検討リスト.xlsx"  # 여기에 URL 목록이 있는 엑셀 파일 경로를 넣어주세요

    # Excel에서 URL 목록 읽기
    url_list = get_urls_from_excel(excel_file_path)

    if not url_list:
        print("URL 목록을 가져올 수 없습니다.")
        return

    # 추출된 URL 목록 확인 (디버깅용)
    print("최종 URL 목록:", url_list)

    # 각 URL에 대해 크롤링 및 엑셀 파일 저장
    for url in url_list:
        print(f"크롤링 중: {url}")
        crawled_data = crawl_course_details(url)
        if crawled_data:
            file_path = r"C:\Users\L1003613honbu048\Documents\業務\メモ\ピクチャー\QC活動\course_details.xlsx"
            save_to_excel([crawled_data], file_path)



if __name__ == "__main__":
    main()