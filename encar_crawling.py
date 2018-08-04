# -*- encoding:utf-8 -*-

import time
import os
from openpyxl import Workbook
from selenium import webdriver
from bs4 import BeautifulSoup

def make_excel(dataList):
    """
        :호출예시 make_excel([ [1,2,3,4], [5,6,7,8] ]) or make_excel(2dArray)
        :param dataList:  [ data1, data2, data3, data4 ] 꼴의 1차원 list를 가지는 2차원 list
        :return: 없
    """
    # === CONFIG
    FILENAME = "엔카.xlsx"

    # === SAVE EXCEL
    wb = Workbook()
    ws1 = wb.worksheets[0]
    header1 = ['제조사', '모델', '세부모델', '등급']
    ws1.column_dimensions['A'].width = 30
    ws1.column_dimensions['B'].width = 30
    ws1.column_dimensions['C'].width = 50
    ws1.column_dimensions['D'].width = 50
    ws1.append(header1)
    # data save

    for data in dataList:
        ws1.append(data)
    # end
    wb.save(FILENAME)


def make_excel_manufacturer(dataList, name):
    """
        :호출예시 make_excel([ [1,2,3,4], [5,6,7,8] ]) or make_excel(2dArray)
        :param dataList:  [ data1, data2, data3, data4 ] 꼴의 1차원 list를 가지는 2차원 list
        :return: 없
    """
    # === CONFIG
    FILENAME = "엔카_" + name + ".xlsx"

    # === SAVE EXCEL
    wb = Workbook()
    ws1 = wb.worksheets[0]
    header1 = ['제조사', '모델', '세부모델', '등급']
    ws1.column_dimensions['A'].width = 30
    ws1.column_dimensions['B'].width = 30
    ws1.column_dimensions['C'].width = 50
    ws1.column_dimensions['D'].width = 50
    ws1.append(header1)
    # data save

    for data in dataList:
        ws1.append(data)
    # end
    wb.save(FILENAME)


def chk_loading():
    bs4 = BeautifulSoup(driver.page_source, 'lxml')
    style_attr = bs4.find('div', class_='case_loading').get('style')
    if style_attr == 'display:none' or style_attr == 'display: none;':
        return True
    else:
        return False

def wait_loading():
    while not chk_loading():
        time.sleep(0.2)


if __name__ == "__main__":
    # ========= BETA SETTING
    """
    now = 1532504559.5943735
    terminTime = now + 60 * 60 * 3
    print("체험판 만료기간 : ", time.ctime(terminTime))
    if time.time() > terminTime:
        print('만료되었습니다.')
        exit(-1)
    """

    # =========
    # SETTING
    setting_list = []
    dir_name_dom = ""
    dir_name_imp = ""
    now_idx = []
    try:
        setting_file = open('setting.ini', 'r', encoding='utf-8')
        setting_file.fileno()
        setting_list = setting_file.readlines()
        print("[COMPLETE] Setting File 확인")

        temp = 0
        for i in setting_list:
            # print(temp, ': ', end='')
            # print(i)
            temp += 1

        # DIRECTORY
        dir_name_dom = setting_list[1].split(':')[1].strip()
        dir_name_imp = setting_list[2].split(':')[1].strip()

        # NOW INDEX
        for i in setting_list[5].split('/'):
            now_idx.append(int(i.strip()))

        setting_file.close()
    except FileNotFoundError:
        print("[ERROR] Setting File 확인 실패")
        exit()

    # DRIVER INITIATE
    driver = webdriver.Chrome('chromedriver.exe')
    driver.maximize_window()

    # VARIABLE
    url_list = [
        'http://www.encar.com/dc/dc_carsearchlist.do?carType=kor&searchType=model&TG.R=A#!',
        'http://www.encar.com/fc/fc_carsearchlist.do?carType=for&searchType=model&TG.R=B#!'
    ]

    url_pivot = now_idx[0]
    depth1_pivot = now_idx[1]
    depth2_pivot = -1
    depth3_pivot = -1

    depth1 = ""
    depth2 = ""
    depth3 = ""
    depth4 = ""

    result = []
    result_temp = []

    cancel_depth1 = '//*[@id="schModelstep"]/div/p/input'
    cancel_depth2 = '//*[@id="schModelstep"]/div/p[2]/input'
    cancel_depth3 = '//*[@id="schModelstep"]/div/p[3]/input'

    current_path = os.getcwd()
    # =========
    # STEP 0 : Directory 생성
    try:
        if not os.path.isdir('./' + dir_name_dom):
            os.mkdir('./' + dir_name_dom)

        if not os.path.isdir('./' + dir_name_imp):
            os.mkdir('./' + dir_name_imp)
        print("[COMPLETE] Directory 생성 완료")
    except:
        print("[ERROR] Directory 생성 실패")
        exit()

    # STEP 1.0 : 국산차 url, 수입차 url 이동
    while url_pivot < 2:

        # STEP 1.0.0 : 현재 Directory 이동
        if url_pivot == 0:
            os.chdir(current_path + '/' + dir_name_dom)
        elif url_pivot == 1:
            os.chdir(current_path + '/' + dir_name_imp)

        driver.get(url_list[url_pivot])
        time.sleep(0.5)

        # STEP 1.1 : 제조사 개수 확인 [ Depth 1 ]
        depth1_cnt = -1
        try:
            bs4 = BeautifulSoup(driver.page_source, 'lxml')
            if url_pivot == 0:
                depth1_cnt = len(bs4.find('div', id='stepManufact').find_all('dd'))
            elif url_pivot == 1:
                depth1_cnt = len(bs4.find('div', id='stepManufact').find('dl', class_='deplist sort_lista').find_all('dd'))
        except:
            print("[ERROR] 제조사 항목 인식 실패")
            driver.implicitly_wait(2)
            time.sleep(1)
            print("......")
            print("[REPAIR] 다시 시도")
            continue

        # STEP 1.2 : 제조사 선택 [ Depth 1 ]
        while depth1_pivot <= depth1_cnt:
            result_temp.clear()
            depth1_dd_x_path = ""
            try:
                if url_pivot == 0:
                    depth1_dd_x_path = '//*[@id="stepManufact"]/dl/dd[{}]'.format(depth1_pivot)
                elif url_pivot == 1:
                    depth1_dd_x_path = '//*[@id="stepManufact"]/dl[2]/dd[{}]'.format(depth1_pivot)
                driver.find_element_by_xpath(depth1_dd_x_path).click()
                time.sleep(0.2)
            except:
                print("[ERROR] 제조사 항목 선택 실패")
                driver.implicitly_wait(2)
                time.sleep(1)
                print("......")
                print("[REPAIR] 다시 시도")
                continue

            wait_loading()

            # STEP 1.3 : 모델 개수 확인 [ Depth 2 ]
            while True:
                try:
                    bs4 = BeautifulSoup(driver.page_source, 'lxml')
                    temp_depth2 = bs4.find('div', id='stepModel').find('dl', class_='deplist sort_lista')
                    case_depth2 = 0
                    if not temp_depth2:
                        temp_depth2 = bs4.find('div', id='stepModel').find('dl', class_='deplist sort_titnon')
                        case_depth2 = 1
                    depth2_cnt = len(temp_depth2.find_all('dd'))
                except:
                    print("[ERROR] 모델 항목 인식 실패")
                    driver.implicitly_wait(2)
                    time.sleep(1)
                    print("......")
                    print("[REPAIR] 다시 시도")
                    continue
                break

            # STEP 1.4 : 모델 선택 [ Depth 2 ]
            depth2_pivot = 1
            while depth2_pivot <= depth2_cnt:
                depth2_dd_x_path = ""
                try:
                    if case_depth2 == 0:
                        depth2_dd_x_path = '//*[@id="stepModel"]/dl[2]/dd[{}]'.format(depth2_pivot)
                    elif case_depth2 == 1:
                        depth2_dd_x_path = '//*[@id="stepModel"]/dl/dd[{}]'.format(depth2_pivot)
                    driver.find_element_by_xpath(depth2_dd_x_path).click()
                    time.sleep(0.2)
                except:
                    print("[ERROR] 모델 항목 선택 실패")
                    driver.implicitly_wait(2)
                    time.sleep(1)
                    print("......")
                    print("[REPAIR] 다시 시도")
                    continue

                wait_loading()

                # STEP 1.5 : 세부 모델 개수 확인 [ Depth 3 ]
                while True:
                    try:
                        bs4 = BeautifulSoup(driver.page_source, 'lxml')
                        depth3_cnt = len(bs4.find('div', id='stepDeModel').find_all('dd'))
                    except:
                        print("[ERROR] 세부 모델 항목 인식 실패")
                        driver.implicitly_wait(2)
                        time.sleep(1)
                        print("......")
                        print("[REPAIR] 다시 시도")
                        continue
                    break

                # STEP 1.6 : 세부 모델 선택 [ Depth 3 ]
                depth3_pivot = 1
                while depth3_pivot <= depth3_cnt:
                    try:
                        depth3_dd_x_path = '//*[@id="stepDeModel"]/dl/dd[{}]'.format(depth3_pivot)
                        driver.find_element_by_xpath(depth3_dd_x_path).click()
                        time.sleep(0.2)
                    except:
                        print("[ERROR] 세부 모델 항목 선택 실패")
                        driver.implicitly_wait(2)
                        time.sleep(1)
                        print("......")
                        print("[REPAIR] 다시 시도")
                        continue

                    wait_loading()

                    # STEP 1.7 : 제조사, 모델, 세부 모델 저장 [ Depth 1, Depth 2, Depth 3 ]
                    while True:
                        try:
                            bs4 = BeautifulSoup(driver.page_source, 'lxml')
                            depth1 = bs4.find('p', class_='choitem step1').find('strong').get_text()
                            depth2 = bs4.find('p', class_='choitem step2').find('strong').get_text()
                            depth3 = bs4.find('p', class_='choitem step3').find('strong').get_text()
                        except:
                            print("[ERROR] 제조사, 모델, 세부 모델 항목 인식 실패")
                            driver.implicitly_wait(2)
                            time.sleep(1)
                            print("......")
                            print("[REPAIR] 다시 시도")
                            continue
                        break

                    # STEP 1.8 : 등급 LIST 확인 [ Depth 4 ]
                    while True:
                        try:
                            depth4_dd_list = []
                            bs4 = BeautifulSoup(driver.page_source, 'lxml')
                            temp = bs4.find('div', id='stepGardeSet')
                            if temp != None:
                                depth4_dd_list = temp.find_all('dd')
                            else:
                                break
                        except:
                            print("[ERROR] 등급 항목 인식 실패")
                            driver.implicitly_wait(2)
                            time.sleep(1)
                            print("......")
                            print("[REPAIR] 다시 시도")
                            continue
                        break

                    # STEP 1.9 : 등급 저장 [ Depth 4 ]
                    temp_result = []
                    if len(depth4_dd_list) != 0:
                        for depth4_dd in depth4_dd_list:
                            temp_result = []
                            depth4 = depth4_dd.find('label').get_text()

                            # STEP 1.10.1 : RESULT LIST에 결과 저장 [ Depth 1, Depth 2, Depth 3, Depth 4 ] (일반 항목)
                            temp_result.append(depth1)
                            temp_result.append(depth2)
                            temp_result.append(depth3)
                            temp_result.append(depth4)
                            result.append(temp_result)
                            result_temp.append(temp_result)
                            print(temp_result)
                    else:
                        # STEP 1.10.2 : RESULT LIST에 결과 저장 [ Depth 1, Depth 2, Depth 3, Depth 4 ] (기타 항목)
                        depth4 = ""
                        temp_result.append(depth1)
                        temp_result.append(depth2)
                        temp_result.append(depth3)
                        temp_result.append(depth4)
                        result.append(temp_result)
                        result_temp.append(temp_result)
                        print(temp_result)

                    # STEP 1.10 : 세부 모델 취소
                    while True:
                        try:
                            driver.find_element_by_xpath(cancel_depth3).click()
                            time.sleep(0.2)
                        except:
                            print("[ERROR] 세부 모델 취소 버튼 인식 실패")
                            driver.implicitly_wait(2)
                            time.sleep(1)
                            print("......")
                            print("[REPAIR] 다시 시도")
                            continue
                        break
                    wait_loading()
                    depth3_pivot += 1

                # STEP 1.11 : 모델 취소
                while True:
                    try:
                        driver.find_element_by_xpath(cancel_depth2).click()
                        time.sleep(0.2)
                    except:
                        print("[ERROR] 모델 취소 버튼 인식 실패")
                        driver.implicitly_wait(2)
                        time.sleep(1)
                        print("......")
                        print("[REPAIR] 다시 시도")
                        continue
                    break
                wait_loading()
                depth2_pivot += 1

            # STEP 1.12 : 제조사 취소
            while True:
                try:
                    driver.find_element_by_xpath(cancel_depth1).click()
                    time.sleep(0.2)
                except:
                    print("[ERROR] 제조사 취소 버튼 인식 실패")
                    driver.implicitly_wait(2)
                    time.sleep(1)
                    print("......")
                    print("[REPAIR] 다시 시도")
                    continue
                break
            wait_loading()
            depth1_pivot += 1

            # STEP 1.13 : 제조사별 Excel 생성
            make_excel_manufacturer(result_temp, depth1)
        url_pivot += 1
        depth1_pivot = 1

    # STEP 1.14 : driver 종료
    driver.quit()

    # STEP 2 : Excel 생성
    os.chdir(current_path)
    make_excel(result)
