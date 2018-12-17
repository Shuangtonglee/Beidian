import time
import os
import datetime
from flask import Flask,send_file,render_template
from flask_apscheduler import APScheduler
from openpyxl import load_workbook,Workbook
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from random import randint

class Config(object):
    JOBS = [
            {
               'id':'job1',
               'func':'beidian:get_data',
               'args': '',
               'trigger': {
                    'type': 'cron',
                    'hour':'0-23',
                    'minute':'00'
                }

             }
        ]

    SCHEDULER_API_ENABLED = True


headers = {'User-Agent':"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1"}
urls = ['https://ibei.cn/1y5uOZ','https://m.beidian.com/detail/detail.html?iid=30409280&shop_id=1480734&utm_source=&offer_code=0']


service_args=[]
service_args.append('--load-images=no')
service_args.append('--disk-cache=yes')
service_args.append('--ignore-ssl-errors=true')



app = Flask(__name__)
scheduler = APScheduler()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download')
def download():
    return send_file('static/sales_data.xlsx',as_attachment=True)

@app.route('/pause')
def pause():
    scheduler.pause_job('job1')
    return render_template('resume.html')


@app.route('/resume')
def resume():
    scheduler.resume_job('job1')
    return render_template('pause.html')


def get_data():
    driver=webdriver.PhantomJS(executable_path='bin/phantomjs',service_args=service_args)
    sales_number_list = []
    for url in urls:
        driver.get(url)
        time.sleep(randint(1,3))
        try:
            element = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.CLASS_NAME,'J_sellerCount')))
        finally:
            sales_number = driver.find_element_by_class_name('J_sellerCount').text
        sales_number_list.append(sales_number)
    driver.quit()
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), sales_number_list)

    if not os.path.exists('static'):
        os.makedirs('static')

    filename_path = 'static/sales_data.xlsx'
    sheet_name = datetime.datetime.now().strftime('%m-%d')
    if not os.path.exists(filename_path):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.append(['时间','销量'])
    else:
        wb = load_workbook(filename=filename_path)
        sheets = wb.get_sheet_names()
        if not sheet_name in sheets:
            ws = wb.create_sheet(sheet_name)
            ws.append(['时间','销量'])
        else:
            ws = wb[sheet_name]


    for sale in sales_number_list:
        ws.append([datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),sale])
    wb.save(filename=filename_path)

def my_listener(event):
    if event.exception:
        print('The job crashed :(')
    else:
        print('The job worked :)')

