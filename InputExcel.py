# InputExcel.py
import openpyxl
import os


def save_to_excel(data, file_path):
    try:
        # 파일 경로가 없으면 새로 생성
        if not os.path.exists(file_path):
            wb = openpyxl.Workbook()  # 새 엑셀 워크북 생성
        else:
            wb = openpyxl.load_workbook(file_path)  # 기존 워크북 열기

        ws = wb.active

        # 첫 번째 행에 제목 추가 (제목이 없으면 추가)
        if ws.max_row == 1:
            ws.append(["コース概要"])

        # 크롤링한 데이터를 엑셀에 추가
        for item in data:
            ws.append([item.get("コース概要", "N/A")])

        # 엑셀 파일 저장
        wb.save(file_path)
        print(f"Excel 파일이 성공적으로 업데이트되었습니다: {file_path}")

    except Exception as e:
        print(f"Excel 업데이트 중 에러 발생: {e}")
