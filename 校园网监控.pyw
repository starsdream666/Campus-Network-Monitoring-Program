import os
import sys
import time
import requests
from bs4 import BeautifulSoup
import subprocess
import threading
import schedule

# 指定网页URL和期望的标题文字
WEB_PAGE_URL = "https://auth.jljy.edu.cn"
EXPECTED_TITLE = "上网登录页"
SCRIPT_PATH = "D:/WindowsGet/login.pyw"

def get_page_title(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        return soup.title.string
    except Exception as e:
        print(f"获取页面标题时出错: {e}")
        return None

def run_script():
    try:
        subprocess.run(["python", SCRIPT_PATH], check=True)
    except subprocess.CalledProcessError as e:
        print(f"脚本执行失败: {e}")

def check_and_run():
    title = get_page_title(WEB_PAGE_URL)
    if title is not None:
        print(f"当前页面标题: {title}")
        if title == EXPECTED_TITLE:
            print("标题匹配，运行脚本...")
            threading.Thread(target=run_script).start()
    else:
        print("无法获取页面标题")

def job():
    check_and_run()

def start_scheduler():
    schedule.every(10).seconds.do(job)  # 每10秒运行一次任务

    while True:
        schedule.run_pending()
        time.sleep(1)

def main():
    try:
        start_scheduler()
    except Exception as e:
        print(f"程序发生异常: {e}")
        # 记录日志等操作

        # 重新启动程序
        python = sys.executable
        os.execl(python, python, *sys.argv)

if __name__ == "__main__":
    main()
