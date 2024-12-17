# Crawling.py
import requests
from bs4 import BeautifulSoup


def crawl_course_details(url):
    # URLからHTMLを修得
    response = requests.get(url)
    if response.status_code != 200:
        print(f"Error: Unable to fetch page. Status code: {response.status_code}")
        return None

    # BeautifulSoupでHTMLパーシング
    soup = BeautifulSoup(response.text, 'html.parser')

    # コース概要 を探してクローリング
    course_outline = soup.find('tr', {'id': 'courseDetailLearnDetail_trOutline'})  # 'コース概要'を含むtrタグ
    if course_outline:
        course_overview = course_outline.find_all('td', {'class': 'BOTTOMLINNER2 border-box-course-detail'})
        if course_overview:
            overview_text = course_overview[0].get_text(strip=True)
            return {'コース概要': overview_text}
        else:
            return {'コース概要': 'N/A'}
    else:
        return {'コース概要': 'N/A'}
