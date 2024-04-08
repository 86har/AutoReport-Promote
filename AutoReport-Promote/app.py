from flask import Flask, render_template, request
import requests
from bs4 import BeautifulSoup
from DrissionPage import ChromiumPage, ChromiumOptions
import time, math, threading, platform, json, pytz, os, re

co = ChromiumOptions().auto_port()
co.set_argument('--window-size=1024,768')
# co.set_argument('--headless')
co.set_user_agent(user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')

try:
    page = ChromiumPage(co)
except Exception as e:
    print(e)

app = Flask(__name__)

page.get('https://www.nnn.ed.nico/my_course')

@app.route('/')
def index():
    return render_template('home.html')

@app.route('/subject-page')
def subject_page():
    subjectIndex = request.args.get('subject')

    class_index = [
        "論理国語",
        "日本史探究",
        "数学Ａ",
        "生物基礎",
        "体育Ⅱ",
        "保健",
        "美術Ⅰ",
        "論理・表現Ⅰ",
        "家庭総合",
        "総合的な探究の時間Ⅱ",
        "特別活動Ⅱ"
    ]
    data = {
        'subjectName': class_index[int(subjectIndex)-1]
    }

    # 外部のWebページを取得する
    url = 'https://www.nnn.ed.nico/courses/1848'
    response = requests.get(url)

    # ページの取得に成功したか確認する
    if response.status_code == 200:
        html = response.text

        # HTMLを解析する
        soup = BeautifulSoup(html, 'html.parser')
        print(soup)
    else:
        print('Failed to fetch the page:', response.status_code)

    return render_template('subject_page.html', data=data)

if __name__ == '__main__':
    app.run(debug=False)
