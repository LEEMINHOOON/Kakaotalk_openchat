import time
import win32con, win32api, win32gui
import requests
from bs4 import BeautifulSoup
from apscheduler.schedulers.background import BackgroundScheduler

# 카톡창 이름, (활성화 상태의 열려있는 창)
kakao_opentalk_name = '아주대학교 디지털미디어학과 공지방'

# AllForYoung 사이트 URL (최신 게시글이 있는 페이지)
target_url = 'https://www.ajou.ac.kr/kr/ajou/notice.do'  # 최신 글 목록 페이지로 교체

# 이전에 확인한 마지막 글 ID 저장
last_post_id = None

# 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, text):
    # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx(hwndMain, None, "RichEdit20W", None)

    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)

# 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# 채팅방 열기
def open_chatroom(chatroom_name):
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx(hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx(hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx(hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx(hwndkakao_edit2_2, None, "Edit", None)

    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)  # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(1)

# AllForYoung 사이트에서 최신 글 확인
def check_new_post():
    global last_post_id
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 '
                      '(KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36'
    }

    res = requests.get(target_url, headers=headers)
    soup = BeautifulSoup(res.content, 'html.parser')

    # 최신 글의 ID와 내용을 파싱 (구조에 따라 수정 필요)
    latest_post = soup.find('div', class_='post')  # 게시글 리스트 구조에 맞게 수정 필요
    latest_post_id = latest_post['data-post-id']  # 게시글 ID를 가져온다고 가정
    latest_post_content = latest_post.get_text()  # 게시글 내용을 가져옴

    # 새로운 글이 확인되면 전송
    if last_post_id is None or latest_post_id != last_post_id:
        last_post_id = latest_post_id  # 최신 글 ID 업데이트
        return latest_post_content

    return None  # 새로운 글이 없으면 None 반환

# 스케줄러 job_1
def job_1():
    p_time_ymd_hms = \
        f"{time.localtime().tm_year}-{time.localtime().tm_mon}-{time.localtime().tm_mday} / " \
        f"{time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"

    open_chatroom(kakao_opentalk_name)  # 채팅방 열기
    new_post_content = check_new_post()  # 새로운 글 확인

    if new_post_content:
        kakao_sendtext(kakao_opentalk_name, f"{p_time_ymd_hms}\n{new_post_content}")  # 메시지 전송, time/새 글 내용

def main():
    sched = BackgroundScheduler()
    sched.start()

    # 매 분 5초마다 job_1 실행
    sched.add_job(job_1, 'cron', second='*/5', id="test_1")

    while True:
        print("실행중.................")
        time.sleep(1)

if __name__ == '__main__':
    main()