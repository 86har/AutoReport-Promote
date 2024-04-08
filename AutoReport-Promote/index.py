from DrissionPage import ChromiumPage, ChromiumOptions
import time, math, threading, platform, json, pytz, os, re

co = ChromiumOptions()
co.set_argument('--window-size=1024,768')
co.set_user_agent(user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')

try:
    page = ChromiumPage(co)
except Exception as e:
    print(e)

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = r"D:\programs\AutoReport-Promote\asetts\AuthenticationInformation\reports-state-of-progress-api-fa1517d1beb9.json"
credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

def changeCellColor(col, row, color):
    request = {
        "updateCells": {
            "rows": [{"values": [{ "userEnteredFormat":{ "backgroundColor": color}}]}],
            "fields": "userEnteredFormat.backgroundColor",
            "start": {
                "sheetId": "0",  # シートのID
                "rowIndex": row,  # 行のインデックス
                "columnIndex": col  # 列のインデックス
            }
        }
    }
    spreadsheet_id = '1QnMkd6GT38IiidmpK3y59lPW5Mn5lKKf1QFbeFs7cho'
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [request]}
    ).execute()

green_color = {
    "red": 0.3,
    "green": 1.0,
    "blue": 0.3
}
red_color = {
    "red": 1.0,
    "green": 0.3,
    "blue": 0.2
}

def hasClassElm(className):
    return page.run_js(f'''
        let className = '{className}';
        className = ' ' + className;
        if(document.querySelectorAll(className.replace(/ /g, '.')).length > 0) return true;
        return false;
    ''')

def whileHasClass(className):
    while(page.run_js(f'''
        let className = '{className}';
        className = ' ' + className;
        if(document.querySelectorAll(className.replace(/ /g, '.')).length > 0) return false;
        return true;
    ''')):
        time.sleep(0.01)

def clickClassElm(className, index = 0, timeout = 10):
    start_time = time.time()
    while True:
        if hasClassElm(className):
            page.run_js(f'''
                let className = '{className}';
                className = ' ' + className;
                document.querySelectorAll(className.replace(/ /g, '.'))[{index}].click();
            ''')
            return True
        if time.time() - start_time >= timeout:
            print()
            return False

def hasIdElm(idName):
    return page.run_js(f'''
        let idName = '{idName}';
        idName = '#' + idName;
        if(document.querySelectorAll(idName).length > 0) return true;
        return false;
    ''')

def clickIdElm(className, index = 0, timeout = 10):
    start_time = time.time()
    while True:
        if hasIdElm(className):
            page.run_js(f'''
                let className = '{className}';
                className = '#' + className;
                document.querySelectorAll(className)[{index}].click();
            ''')
            return True
        if time.time() - start_time >= timeout:
            print()
            return False
        
        time.sleep(0.01)

page.get('https://www.nnn.ed.nico/my_course')

time.sleep(0.5)

if hasClassElm('sc-aXZVg tTAOW sc-13j7nb-0 sc-1204fnl-0 klVKWm eFvVSa'):
    clickClassElm('sc-aXZVg tTAOW sc-13j7nb-0 sc-1204fnl-0 klVKWm eFvVSa')
    time.sleep(0.5)
    clickClassElm('sc-aXZVg dLyEch sc-13j7nb-0 sc-8ve2o9-3 sc-8ve2o9-4 klVKWm dvROaM cwkWKR')
    page.ele('#oauth_identifier_loginId', timeout=10).input('23N1101007')
    page.ele('#oauth_identifier_password', timeout=10).input('8610haruhAru')
    time.sleep(0.5)
    clickIdElm('oauth_identifier_')
time.sleep(2)

while(page.run_js(f'return document.querySelector(".sc-1x2znj8-0.gfxzXN") === null;')):
    time.sleep(0.01)

reportNums = page.run_js('return document.querySelector(".sc-1x2znj8-0.gfxzXN").children.length')
for i in range(reportNums):
    while(page.run_js(f'return document.querySelector(".sc-1x2znj8-0.gfxzXN") === null;')):
        time.sleep(0.01)

    restReportNums = page.run_js(f'''
        let figureElm=document.querySelector(".sc-1x2znj8-0.gfxzXN").children[{i}].getElementsByTagName("figure")[0].getAttribute("aria-label").match(/\d+/g);
        return figureElm[0]-figureElm[1];
    ''')

    if(restReportNums > 0):
        page.run_js(f'document.querySelector(".sc-1x2znj8-0.gfxzXN").children[{i}].children[0].click()')
        while(page.run_js(f'return document.querySelector(".sc-1bplusx-0.fdRxUg") === null;')):
            time.sleep(0.01)
        time.sleep(0.05)

        restReportListNums = page.run_js('return document.querySelector(".sc-1bplusx-0.fdRxUg").children.length;')
        watchedCount = 0

        for j in range(restReportListNums):
            className = page.run_js(f'return document.querySelector(".sc-1bplusx-0.fdRxUg").children[{j}].querySelector(".sc-aXZVg.dKubqp.sc-ni8l2q-3.kwUDW").innerText.split(":")[0]')
            reportsNumber = page.run_js(f'return document.querySelector(".sc-1bplusx-0.fdRxUg").children[{j}].querySelector(".sc-aXZVg.dKubqp.sc-ni8l2q-2.bIeKIR").innerText.match(/第(\d+)回/)[1]')

            page.run_js(f'document.querySelector(".sc-1bplusx-0.fdRxUg").children[{j}].children[0].click()')

            attempts = 0
            while(page.run_js(f'return document.querySelector(".sc-aXZVg.tTAOW.sc-13j7nb-0.sc-18fty53-0.sc-nwys29-0.klVKWm.ioXRWf.lePrcr") === null;')):
                time.sleep(0.01)
                if attempts>=20: break
                attempts+=1
            time.sleep(0.05)

            try:
                page.run_js('document.querySelector(".sc-aXZVg.tTAOW.sc-13j7nb-0.sc-18fty53-0.sc-nwys29-0.klVKWm.ioXRWf.lePrcr").click();')
            except:
                pass

            while(page.run_js(f'return document.querySelector(".sc-aXZVg.sc-gEvEer.sc-l5r9s4-0.dKubqp.fteAEG.bKupGM") === null;')):
                time.sleep(0.01)

            isWatched = page.run_js(f''' 
                let nextReportElm = document.querySelector(".sc-aXZVg.sc-gEvEer.sc-l5r9s4-0.dKubqp.fteAEG.bKupGM").querySelector(".sc-aXZVg.sc-gEvEer.hYNtMZ.fteAEG.sc-1otp79h-0.sc-35qwhb-0.evJGlU.hoWVG");
                let lessonTitle = document.querySelector(".sc-1h5ye17-0.ipaxCh").innerText;
                        
                let isWatched = false;
                                          
                if( nextReportElm == null ){{
                    console.log("この教材は視聴済みです");
                    isWatched = true;
                    document.querySelector(".sc-1mr8gis-5.ehXNoQ").click();
                }}else{{
                    if( nextReportElm.parentNode.previousElementSibling.querySelectorAll(".sc-x54faw-0.fraACz").length > 0){{
                        console.log("この教材は未視聴です。");
                        nextReportElm.parentNode.previousElementSibling.children[0].click();
                    }}else{{
                        console.log("この教材は視聴済みです");
                        isWatched = true;
                        document.querySelector(".sc-1mr8gis-5.ehXNoQ").click();
                    }} 
                }}         
                
                return isWatched;
            ''')

            if(isWatched):
                watchedCount += 1
            else:
                while(1):
                    while(page.run_js(f'''
                                      try{{
                                        document.getElementsByTagName("iframe")[0].contentWindow.document.querySelectorAll("#video-player")[0].ended;
                                      }}catch(e){{
                                        return true;
                                      }}
                                      return false;
                                      ''')):
                        time.sleep(0.01)
                    time.sleep(0.05)

                    ended = page.run_js('return document.getElementsByTagName("iframe")[0].contentWindow.document.querySelectorAll("#video-player")[0].ended;')

                    if ended:
                        time.sleep(4)
                        continuation = page.run_js('''
                            const videoLists = document.getElementsByClassName("sc-aXZVg sc-gEvEer sc-l5r9s4-0 dKubqp fteAEG bKupGM")[0].children;
                            let videoIdentification = false;
                
                            let targetNumber;
                            let targetType;
                            for(let i = 0; i<videoLists.length; i++){
                                let videoElm = videoLists[i];
                                if(videoIdentification){
                                    if(videoElm.querySelectorAll('.sc-x54faw-0.fraACz').length > 0){
                                        targetNumber=i;
                                        targetType="video";
                                        break;
                                    }else if(videoElm.querySelectorAll('.sc-x54faw-0.fsXBnY').length > 0){
                                        targetNumber = i;
                                        targetType = "test";
                                        break;
                                    }
                                }
                                if(videoElm.querySelectorAll('.sc-aXZVg.sc-gEvEer.sc-lcfvsp-10.dKubqp.fteAEG.zixPn').length > 0){
                                    videoIdentification = true;
                                }
                            };
                                    
                            if(targetType == "video"){
                                videoLists[targetNumber].children[0].click();
                                return true;
                            }else{
                                document.querySelectorAll(".sc-1mr8gis-1.eCJxrg")[1].click();
                                return false;
                            }
                        ''')
                        if not continuation:
                            break
                        else:
                            time.sleep(1.5)
            class_index = {
                "論理国語": 1,
                "日本史探究": 2,
                "数学Ａ": 3,
                "生物基礎": 4,
                "体育Ⅱ": 5,
                "保健": 6,
                "美術Ⅰ": 7,
                "論理・表現Ⅰ": 8,
                "家庭総合": 9,
                "総合的な探究の時間Ⅱ": 10,
                "特別活動Ⅱ": 11
            }
            changeCellColor(reportsNumber, class_index.get(className.strip()), green_color)
        clickClassElm("sc-1mr8gis-1 eCJxrg")

page.quit()