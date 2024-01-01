
import tkinter.ttk as ttk
import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import tkinter.messagebox as msgbox
from tkinter import * # __all__
from tkinter import filedialog
import sys, os
import numpy as np
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
from pandas import ExcelWriter
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import shutil  # shutil 모듈 추가
import random
from openpyxl.styles import Border, Side
import subprocess
import win32com.client as win32
import pyperclip
from PIL import Image

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

root = TkinterDnD.Tk()
root.title("Automatic Conversion Program_RPAv.0.1_iron")
root.geometry("550x500")

# 엑셀파일 추가: 첫번째, 버튼클릭 후 경로에서 추가 
def sa_add_file():
    sa_files = filedialog.askopenfilenames(title="업데이트한 엑셀 파일을 추가해 주세요~!", \
        filetypes=(("엑셀 파일", "*.xlsx"), ("모든 파일", "*.*")), \
        initialdir=os.path)
        # 최초에 사용자가 지정한 경로를 보여줌, initialdir=r"C:/")
    
    # 사용자가 선택한 파일 목록
    for file in sa_files:
        sa_list_file.insert(END, file)
    # print(list_file.info)

# 엑셀파일 추가: 두번째, 드래그앤드롭
def dragDrop(event):
    # 파일 경로를 줄 단위로 나누고 각 파일 경로에 따옴표를 추가하여 리스트에 추가
    sa_files = [file.strip() for file in event.data.split("\n") if file.strip()]

    ## ★방법1(1줄에 통합변경)★ 중괄호 제거와 '{ }'를 기준으로 각각 다른 셀로 구분하는 방법 
    sa_files = [file for files in [file.strip('{}').split('} {') for file in sa_files] for file in files]
    for file in sa_files:       # for문도 간단한 편
        # 중복 추가 방지
        if file not in sa_list_file.get(0, tk.END):
            sa_list_file.insert(tk.END, file)

    # ## ★방법2(각 단계별 분리)★ 중괄호 제거와 '{ }'를 기준으로 각각 다른 셀로 구분하는 방법
    # # 중괄호 제거(sa_files의 처음과 끝의 중괄호만 제거)
    # sa_files = [file.strip('{}') for file in sa_files]
    # # '} {'를 기준으로 각각 다른 셀로 구분
    # sa_files = [file.split('} {') for file in sa_files]
    # # 중첩리스트 생성되었고, 이를 반복해서 리스트에 load(방법1 for문보다 복잡)
    # for files in sa_files:  # 수정: 'file'을 'files'로 변경
    #     for file in files:  # 수정: 중첩된 리스트에서 각 파일에 대해 반복
    #         # 중복 추가 방지
    #         if file not in sa_list_file.get(0, tk.END):
    #             sa_list_file.insert(tk.END, file)

# 엑셀파일 선택 삭제
def sa_del_file():
    #print(list_file.curselection())
    for index in reversed(sa_list_file.curselection()):
        sa_list_file.delete(index)


# 결과파일 저장 (폴더)
def browse_dest_path():
    folder_selected = filedialog.askdirectory()
    if folder_selected == "": # 사용자가 취소를 누를 때
        print("폴더 선택 취소")
        return
    # 'images' 폴더 경로 생성
    images_folder_path = os.path.join(folder_selected, 'images')
    # 'images' 폴더 생성
    os.makedirs(images_folder_path, exist_ok=True)
    #print(folder_selected)
    txt_dest_path.delete(0, END)
    txt_dest_path.insert(0, folder_selected)

def chart_ext():
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True  # Excel 창을 보이게 설정
    for chart_name in chart_names:
        try:
            # 원하는 차트 활성화
            chart = sheet.ChartObjects(chart_name).Activate()

            # 차트를 이미지로 저장
            images_folder_path = os.path.join(txt_dest_path.get(), 'images')
            chart_image_path = os.path.join(images_folder_path, f"{chart_name}.jpg")
            excel_app.ActiveChart.Export(chart_image_path, FilterName="JPG")

    #         print(f"차트 '{chart_name}'를 '{chart_image_path}'에 저장했습니다.")
        except Exception as e:
            print(f"차트 '{chart_name}'를 저장하는 도중 에러 발생: {e}")

## 업황지수 int, p단위 함수
def int_judge_num(theNumber):
    result = ""
    if round(theNumber) > 0:
        result = str(theNumber) + "p 상승"
    elif round(theNumber) < 0:
        result = str(abs(int(theNumber))) + "p 하락"  # 문자열로 변환 후 연결
    else:
        result = "변동 없음"
    return result

## 업황지수 float 소수첫째자리, p단위 함수
def float_judge_num(theNumber):
    result = ""
    if round(theNumber, 1) > 0:
        result = str(round(float(theNumber), 1)) + "p 상승"  # round 함수 수정
    elif round(theNumber, 1) < 0:
        result = str(round(abs(float(theNumber)), 1)) + "p 하락"  # 문자열로 변환 후 연결
    else:
        result = "변동 없음"
    return result

## 수출입 백만불 단위, 억원 단위 함수
def int_billion_num(theNumber):
    result = ""
    if int(theNumber) > 0:
        result = "{:,}".format(round(theNumber, 1)) + "억원 증가"
    elif int(theNumber) < 0:
        result = "{:,}".format(round(abs(theNumber), 1)) + "억원 감소"  # 문자열로 변환 후 연결
    else:
        result = "변동 없음"
    return result
      

##==== 그림파일 폴더('txt_dest_path.get(), 'image')에 있는 그림을 한글에 붙여넣기
# 그림 이미지 척도_mm로 표준화
def mm_to_pixels(mm, dpi=96):
    return int((mm / 25.4) * dpi)

def resize_image(image_path, width_mm, height_mm, dpi=96):
    width_pixels = mm_to_pixels(width_mm, dpi)
    height_pixels = mm_to_pixels(height_mm, dpi)

    img = Image.open(image_path)
    resized_img = img.resize((width_pixels, height_pixels), Image.LANCZOS)
    resized_img.save(image_path)

def hwp에_이미지_붙여넣기(hwp, 필드_이름, 이미지_경로):
    hwp.MoveToField(필드_이름)
    hwp.InsertPicture(이미지_경로, True)
    hwp.Run("Cancel")
    hwp.Run("Cancel")

# 시작(파일작성 버튼 클릭시 동작)
def start():
    if sa_list_file.size() != 1:
        msgbox.showwarning("경고", "'사전 협의된 업데이트한 엑셀파일'이 1개만 있어야 합니다!")
        return

    # # 저장 경로 확인
    if len(txt_dest_path.get()) == 0:
        msgbox.showwarning("경고", "저장할 폴더를 선택하세요")
        return

    ##================= Excel 어플리케이션 시작 및 시트별 차트 모두 가져오기
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True  # Excel 창을 보이게 설정

    # Excel 파일 열기
    workbook = excel_app.Workbooks.Open(sa_list_file.get(0))

    # 12개 차트 가져오기: 경기동향_기업경기 시트 
    global sheet, chart_names
    sheet_name = "경기동향_기업경기"
    sheet = workbook.Sheets(sheet_name)
    # 원하는 차트 이름 리스트_사전작업 필요_차트 선택 후 이름 입력
    chart_names = ["chart1", "chart2", "chart3", "chart4", "chart5", "chart6", "chart7", "chart8", "chart9", "chart10", "chart11", "chart12"]
    chart_ext()
    list_columns = ['A', 'B', 'C', 'D', 'E', 'F']
    column_range = 'A:F'
    df1 = pd.read_excel(sa_list_file.get(0), sheet_name=sheet_name, skiprows=3, usecols=column_range, names=list_columns, header=None)
    df1 = df1.fillna('')

    # 6개 차트 가져오기: 경기동향_소비 시트 
    sheet_name = "경기동향_소비"
    sheet = workbook.Sheets(sheet_name)
    # 원하는 차트 이름 리스트_사전작업 필요_차트 선택 후 이름 입력
    chart_names = ["chart_1", "chart_2", "chart_3", "chart_4", "chart_5", "chart_6"]
    chart_ext()
    list_columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    column_range = 'A:H'
    df2 = pd.read_excel(sa_list_file.get(0), sheet_name=sheet_name, skiprows=3, usecols=column_range, names=list_columns, header=None)
    df2 = df2.fillna('')  

    # 6개 차트 가져오기: 산업동향_제조,건설 시트 
    sheet_name = "산업동향_제조,건설"
    sheet = workbook.Sheets(sheet_name)
    # 원하는 차트 이름 리스트_사전작업 필요_차트 선택 후 이름 입력
    chart_names = ["chart_a", "chart_b", "chart_c", "chart_d", "chart_e", "chart_f"]
    chart_ext()
    df3 = pd.read_excel(sa_list_file.get(0), sheet_name=sheet_name, skiprows=4, usecols=column_range, names=list_columns, header=None)
    df3 = df3.fillna('')

    # 6개 차트 가져오기: 산업동향_서비스업 시트 
    sheet_name = "산업동향_서비스업"
    sheet = workbook.Sheets(sheet_name)
    list_columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
    column_range = 'A:M'
    df4 = pd.read_excel(sa_list_file.get(0), sheet_name=sheet_name, skiprows=4, usecols=column_range, names=list_columns, header=None)
    df4 = df4.fillna('') 

    # 원하는 차트 이름 리스트_사전작업 필요_차트 선택 후 이름 입력
    s_chart2 = df4.iloc[0, 12]
    s_chart4 = df4.iloc[17, 12]
    s_chart6 = df4.iloc[33, 12]
    chart_names = ["s_chart1", s_chart2, "s_chart3", s_chart4, "s_chart5", s_chart6]

    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True  # Excel 창을 보이게 설정

    for i, chart_name in enumerate(chart_names):
        try:
            # 원하는 차트 활성화
            chart = sheet.ChartObjects(chart_name).Activate()

            # 차트를 이미지로 저장
            images_folder_path = os.path.join(txt_dest_path.get(), 'images')
            chart_image_path = os.path.join(images_folder_path, f"s_chart_{i+1}.jpg")
            excel_app.ActiveChart.Export(chart_image_path, FilterName="JPG")

            # print(f"차트 '{chart_name}'를 '{chart_image_path}'에 저장했습니다.")
        except Exception as e:
            print(f"차트 '{chart_name}'를 저장하는 도중 에러 발생: {e}")
    

    # 3개 차트 가져오기: 산업동향_수출입 시트 
    sheet_name = "산업동향_수출입"
    sheet = workbook.Sheets(sheet_name)
    # 원하는 차트 이름 리스트_사전작업 필요_차트 선택 후 이름 입력
    chart_names = ["i_chart1", "i_chart2", "i_chart3"]
    chart_ext()
    list_columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    column_range = 'A:Z'
    df5 = pd.read_excel(sa_list_file.get(0), sheet_name=sheet_name, skiprows=4, usecols=column_range, names=list_columns, header=None)
    df5 = df5.fillna('')  

    # 3개 차트 가져오기: 산업동향_자동차,조선 시트 
    sheet_name = "산업동향_자동차,조선"
    sheet = workbook.Sheets(sheet_name)
    # 원하는 차트 이름 리스트_사전작업 필요_차트 선택 후 이름 입력
    chart_names = ["car_chart1", "car_chart2", "car_chart3"]
    chart_ext()
    list_columns = ['A', 'B', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB']
    column_range = 'A:B,P:AB'
    df6 = pd.read_excel(sa_list_file.get(0), sheet_name=sheet_name, skiprows=4, usecols=column_range, names=list_columns, header=None)
    df6 = df6.fillna('') 

    # 차트없이 자료만 가져오기: 제조_세부 시트
    sheet_name = "제조_세부"
    list_columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    column_range = 'A:H'
    df7 = pd.read_excel(sa_list_file.get(0), sheet_name=sheet_name, usecols=column_range, names=list_columns, header=None)
    df7 = df7.fillna('')

    # Excel 어플리케이션 종료 (저장 여부 묻지 않고 종료)
    excel_app.DisplayAlerts = False
    excel_app.Quit()    


    ### >>>>>>>여기에 데이터프레임에서 값 가져오는 구문 작성 후 붙여넣고
    당월 = "'" + str(df1.iloc[13,0])[2:] + "월"
    전월 = "'" + str(df1.iloc[12,0])[2:] + "월"
    전망월 = "'" + str(df1.iloc[33,0])[2:] + "월"
    당분기 = "'" + str(df4.iloc[8,0][2:8])  # str 자료의 슬라이싱으로 분기값 추출
    ## 부산
    b_m_n = df1.iloc[13, 2]
    theNumber = df1.iloc[13, 3]
    b_m_n_y = int_judge_num(theNumber)
    b_t_n = df1.iloc[13, 4]
    theNumber = df1.iloc[13, 5]
    b_t_n_y = int_judge_num(theNumber)
    b_m_p = df3.iloc[12, 2]
    theNumber = df3.iloc[12, 3]
    b_m_p_y = float_judge_num(theNumber)
    b_m_f = df1.iloc[33, 2]
    theNumber = df1.iloc[33, 3]
    b_m_f_y = int_judge_num(theNumber)
    b_t_f = df1.iloc[33, 4]
    theNumber = df1.iloc[33, 5]
    b_t_f_y = int_judge_num(theNumber)
    b_s_p = df2.iloc[12, 2]
    theNumber = df2.iloc[12, 3]
    b_s_p_y = float_judge_num(theNumber)
    b_p_n = df2.iloc[31, 2]
    theNumber = df2.iloc[31, 3]
    b_p_n_y = float_judge_num(theNumber)
    theNumber = df2.iloc[12, 5]
    b_d_p_y = float_judge_num(theNumber)
    theNumber = df2.iloc[12, 7]
    b_l_p_y = float_judge_num(theNumber)
    b_m_sep1 = df7.iloc[2, 7]
    b_m_sep2 = df7.iloc[3, 7]
    b_m_determin = df7.iloc[0, 0]
    b_m_sep1_1 = df7.iloc[16, 1]
    b_m_sep1_2 = df7.iloc[16, 2]
    b_m_sep2_1 = df7.iloc[17, 1]
    b_m_sep2_2 = df7.iloc[17, 2]
    theNumber = round(df3.iloc[31, 2])
    b_c_p = "{:,}".format(theNumber)
    theNumber = round(df3.iloc[31, 3])
    b_c_p_y = int_billion_num(theNumber)
    theNumber = round(df3.iloc[31, 5])
    b_c_p_const = int_billion_num(theNumber)
    theNumber = round(df3.iloc[31, 7])
    b_c_p_ground = int_billion_num(theNumber)
    b_ss_q = df4.iloc[8, 2]
    theNumber = df4.iloc[8, 3]
    b_ss_q_y = float_judge_num(theNumber)
    b_ss_q_sublime = df4.iloc[12, 8]
    b_i_p = "{:,}".format(round(df5.iloc[28, 2]))
    b_i_p_z = "{:,}".format(round(abs((df5.iloc[28, 3]))))
    b_i_p_y = df5.iloc[28, 6]
    b_i_p_d = df5.iloc[29, 6]
    b_e_p = "{:,}".format(round(df5.iloc[28, 4]))
    b_e_p_z = "{:,}".format(round(abs((df5.iloc[28, 5]))))
    b_e_p_y = df5.iloc[28, 7]
    b_e_p_d = df5.iloc[29, 7]
    b_carship_p_d = df6.iloc[14, 1]
    b_car_p_z = df6.iloc[13, 2]
    b_car_n_z = df6.iloc[13, 3]
    b_ship_p_z = df6.iloc[14, 2]
    b_ship_n_z = df6.iloc[14, 3]
    ## 울산
    u_m_n = df1.iloc[52, 2]
    theNumber = df1.iloc[52, 3]
    u_m_n_y = int_judge_num(theNumber)
    u_t_n = df1.iloc[52, 4]
    theNumber = df1.iloc[52, 5]
    u_t_n_y = int_judge_num(theNumber)
    u_m_p = df3.iloc[48, 2]
    theNumber = df3.iloc[48, 3]
    u_m_p_y = float_judge_num(theNumber)
    u_m_f = df1.iloc[72, 2]
    theNumber = df1.iloc[72, 3]
    u_m_f_y = int_judge_num(theNumber)
    u_t_f = df1.iloc[72, 4]
    theNumber = df1.iloc[72, 5]
    u_t_f_y = int_judge_num(theNumber)
    u_s_p = df2.iloc[47, 2]
    theNumber = df2.iloc[47, 3]
    u_s_p_y = float_judge_num(theNumber)
    u_p_n = df2.iloc[66, 2]
    theNumber = df2.iloc[66, 3]
    u_p_n_y = float_judge_num(theNumber)
    theNumber = df2.iloc[47, 5]
    u_d_p_y = float_judge_num(theNumber)
    theNumber = df2.iloc[47, 7]
    u_l_p_y = float_judge_num(theNumber)
    u_m_sep1 = df7.iloc[7, 7]
    u_m_sep2 = df7.iloc[8, 7]
    u_m_determin = df7.iloc[5, 0]
    u_m_sep1_1 = df7.iloc[19, 1]
    u_m_sep1_2 = df7.iloc[19, 2]
    u_m_sep2_1 = df7.iloc[20, 1]
    u_m_sep2_2 = df7.iloc[20, 2]
    theNumber = round(df3.iloc[66, 2])
    u_c_p = "{:,}".format(theNumber)
    theNumber = round(df3.iloc[66, 3])
    u_c_p_y = int_billion_num(theNumber)
    theNumber = round(df3.iloc[66, 5])
    u_c_p_const = int_billion_num(theNumber)
    theNumber = round(df3.iloc[66, 7])
    u_c_p_ground = int_billion_num(theNumber)
    u_ss_q = df4.iloc[25, 2]
    theNumber = df4.iloc[25, 3]
    u_ss_q_y = float_judge_num(theNumber)
    u_ss_q_sublime = df4.iloc[29, 8]
    u_i_p = "{:,}".format(round(df5.iloc[28, 11]))
    u_i_p_z = "{:,}".format(round(abs((df5.iloc[28, 12]))))
    u_i_p_y = df5.iloc[28, 15]
    u_i_p_d = df5.iloc[29, 15]
    u_e_p = "{:,}".format(round(df5.iloc[28, 13]))
    u_e_p_z = "{:,}".format(round(abs((df5.iloc[28, 14]))))
    u_e_p_y = df5.iloc[28, 16]
    u_e_p_d = df5.iloc[29, 16]
    u_carship_p_d = df6.iloc[18, 1]
    u_car_p_z = df6.iloc[17, 2]
    u_car_n_z = df6.iloc[17, 3]
    u_ship_p_z = df6.iloc[18, 2]
    u_ship_n_z = df6.iloc[18, 3]
    u_ing_p_z = df6.iloc[19, 2]
    u_ing_n_z = df6.iloc[19, 3]
    #경남
    k_m_n = df1.iloc[92, 2]
    theNumber = df1.iloc[92, 3]
    k_m_n_y = int_judge_num(theNumber)
    k_t_n = df1.iloc[92, 4]
    theNumber = df1.iloc[92, 5]
    k_t_n_y = int_judge_num(theNumber)
    k_m_p = df3.iloc[84, 2]
    theNumber = df3.iloc[84, 3]
    k_m_p_y = float_judge_num(theNumber)
    k_m_f = df1.iloc[112, 2]
    theNumber = df1.iloc[112, 3]
    k_m_f_y = int_judge_num(theNumber)
    k_t_f = df1.iloc[112, 4]
    theNumber = df1.iloc[112, 5]
    k_t_f_y = int_judge_num(theNumber)
    k_s_p = df2.iloc[84, 2]
    theNumber = df2.iloc[84, 3]
    k_s_p_y = float_judge_num(theNumber)
    k_p_n = df2.iloc[104, 2]
    theNumber = df2.iloc[104, 3]
    k_p_n_y = float_judge_num(theNumber)
    theNumber = df2.iloc[84, 5]
    k_d_p_y = float_judge_num(theNumber)
    theNumber = df2.iloc[84, 7]
    k_l_p_y = float_judge_num(theNumber)
    k_m_sep1 = df7.iloc[12, 7]
    k_m_sep2 = df7.iloc[13, 7]
    k_m_determin = df7.iloc[10, 0]
    k_m_sep1_1 = df7.iloc[22, 1]
    k_m_sep1_2 = df7.iloc[22, 2]   
    k_m_sep2_1 = df7.iloc[23, 1]
    k_m_sep2_2 = df7.iloc[23, 2]
    theNumber = round(df3.iloc[103, 2])
    k_c_p = "{:,}".format(theNumber)
    theNumber = round(df3.iloc[103, 3])
    k_c_p_y = int_billion_num(theNumber)
    theNumber = round(df3.iloc[103, 5])
    k_c_p_const = int_billion_num(theNumber)
    theNumber = round(df3.iloc[103, 7])
    k_c_p_ground = int_billion_num(theNumber)
    k_ss_q = df4.iloc[41, 2]
    theNumber = df4.iloc[41, 3]
    k_ss_q_y = float_judge_num(theNumber)
    k_ss_q_sublime = df4.iloc[45, 8]
    k_i_p = "{:,}".format(round(df5.iloc[28, 20]))
    k_i_p_z = "{:,}".format(round(abs((df5.iloc[28, 21]))))
    k_i_p_y = df5.iloc[28, 24]
    k_i_p_d = df5.iloc[29, 24]
    k_e_p = "{:,}".format(round(df5.iloc[28, 22]))
    k_e_p_z = "{:,}".format(round(abs((df5.iloc[28, 23]))))
    k_e_p_y = df5.iloc[28, 25]
    k_e_p_d = df5.iloc[29, 25]
    k_carship_p_d = df6.iloc[23, 1]
    k_car_p_z = df6.iloc[22, 2]
    k_car_n_z = df6.iloc[22, 3]
    k_ship_p_z = df6.iloc[23, 2]
    k_ship_n_z = df6.iloc[23, 3]  
    b_ss_n = df4.iloc[0,12][2:]  # 나중에 추가 3개
    u_ss_n = df4.iloc[17,12][2:]
    k_ss_n = df4.iloc[33,12][2:]  ### 데이터 작업 여기까지 끝

    

    #####======================  한글 보고서 작성 =========================================
                  
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    # hwp.XHwpWindows.Item(0).Visible = True

    # 개발문서를 토대로 한글 객체 생성
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

    # 보안모듈 적용
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

    # HWP 파일 경로
    hwp_file_path = resource_path('data/부울경 경제동향(제출)_자동화용.hwpx')

    # 문서 열기
    hwp.Open(hwp_file_path)
    dict_index = {'당월' : 당월, '전월' : 전월, '전망월' : 전망월, '당분기' : 당분기, 'b_m_n' : b_m_n, 'b_m_n_y' : b_m_n_y, 'b_t_n' : b_t_n, 
                'b_t_n_y' : b_t_n_y, 'b_m_p' : b_m_p, 'b_m_p_y' : b_m_p_y, 'b_m_f' : b_m_f, 'b_m_f_y' : b_m_f_y, 'b_t_f' : b_t_f, 
                'b_t_f_y' : b_t_f_y, 'b_s_p' : b_s_p, 'b_s_p_y' : b_s_p_y, 'b_p_n' : b_p_n, 'b_p_n_y' : b_p_n_y, 'b_d_p_y' : b_d_p_y, 
                'b_l_p_y' : b_l_p_y, 'b_m_sep1' : b_m_sep1, 'b_m_sep2' : b_m_sep2, 'b_m_determin' : b_m_determin, 'b_m_sep1_1' : b_m_sep1_1,
                'b_m_sep1_2' : b_m_sep1_2, 'b_m_sep2_1' : b_m_sep2_1, 'b_m_sep2_2' : b_m_sep2_2, 'b_c_p' : b_c_p, 'b_c_p_y' : b_c_p_y, 
                'b_c_p_const' : b_c_p_const, 'b_c_p_ground' : b_c_p_ground, 'b_ss_q' : b_ss_q, 'b_ss_q_y' : b_ss_q_y, 
                'b_ss_q_sublime' : b_ss_q_sublime, 'b_i_p' : b_i_p, 'b_i_p_z' : b_i_p_z, 'b_i_p_y' : b_i_p_y, 'b_i_p_d' : b_i_p_d, 
                'b_e_p' : b_e_p, 'b_e_p_z' : b_e_p_z, 'b_e_p_y' : b_e_p_y, 'b_e_p_d' : b_e_p_d, 'b_carship_p_d' : b_carship_p_d, 
                'b_car_p_z' : b_car_p_z, 'b_car_n_z' : b_car_n_z, 'b_ship_p_z' : b_ship_p_z, 'b_ship_n_z' : b_ship_n_z, 'u_m_n' : u_m_n, 
                'u_m_n_y' : u_m_n_y, 'u_t_n' : u_t_n, 'u_t_n_y' : u_t_n_y, 'u_m_p' : u_m_p, 'u_m_p_y' : u_m_p_y, 'u_m_f' : u_m_f, 
                'u_m_f_y' : u_m_f_y, 'u_t_f' : u_t_f, 'u_t_f_y' : u_t_f_y, 'u_s_p' : u_s_p, 'u_s_p_y' : u_s_p_y, 'u_p_n' : u_p_n, 
                'u_p_n_y' : u_p_n_y, 'u_d_p_y' : u_d_p_y, 'u_l_p_y' : u_l_p_y, 'u_m_sep1' : u_m_sep1, 'u_m_sep2' : u_m_sep2, 
                'u_m_determin' : u_m_determin, 'u_m_sep1_1' : u_m_sep1_1, 'u_m_sep1_2' : u_m_sep1_2, 'u_m_sep2_1' : u_m_sep2_1, 
                'u_m_sep2_2' : u_m_sep2_2, 'u_c_p' : u_c_p, 'u_c_p_y' : u_c_p_y, 'u_c_p_const' : u_c_p_const, 'u_c_p_ground' : u_c_p_ground, 
                'u_ss_q' : u_ss_q, 'u_ss_q_y' : u_ss_q_y, 'u_ss_q_sublime' : u_ss_q_sublime, 'u_i_p' : u_i_p, 'u_i_p_z' : u_i_p_z, 
                'u_i_p_y' : u_i_p_y, 'u_i_p_d' : u_i_p_d, 'u_e_p' : u_e_p, 'u_e_p_z' : u_e_p_z, 'u_e_p_y' : u_e_p_y, 'u_e_p_d' : u_e_p_d, 
                'u_carship_p_d' : u_carship_p_d, 'u_car_p_z' : u_car_p_z, 'u_car_n_z' : u_car_n_z, 'u_ship_p_z' : u_ship_p_z, 
                'u_ship_n_z' : u_ship_n_z, 'u_ing_p_z' : u_ing_p_z, 'u_ing_n_z' : u_ing_n_z, 'k_m_n' : k_m_n, 'k_m_n_y' : k_m_n_y, 
                'k_t_n' : k_t_n, 'k_t_n_y' : k_t_n_y, 'k_m_p' : k_m_p, 'k_m_p_y' : k_m_p_y, 'k_m_f' : k_m_f, 'k_m_f_y' : k_m_f_y, 
                'k_t_f' : k_t_f, 'k_t_f_y' : k_t_f_y, 'k_s_p' : k_s_p, 'k_s_p_y' : k_s_p_y, 'k_p_n' : k_p_n, 'k_p_n_y' : k_p_n_y, 
                'k_d_p_y' : k_d_p_y, 'k_l_p_y' : k_l_p_y, 'k_m_sep1' : k_m_sep1, 'k_m_sep2' : k_m_sep2, 'k_m_determin' : k_m_determin, 
                'k_m_sep1_1' : k_m_sep1_1, 'k_m_sep1_2' : k_m_sep1_2, 'k_m_sep2_1' : k_m_sep2_1, 'k_m_sep2_2' : k_m_sep2_2, 'k_c_p' : k_c_p, 
                'k_c_p_y' : k_c_p_y, 'k_c_p_const' : k_c_p_const, 'k_c_p_ground' : k_c_p_ground, 'k_ss_q' : k_ss_q, 'k_ss_q_y' : k_ss_q_y, 
                'k_ss_q_sublime' : k_ss_q_sublime, 'k_i_p' : k_i_p, 'k_i_p_z' : k_i_p_z, 'k_i_p_y' : k_i_p_y, 'k_i_p_d' : k_i_p_d, 
                'k_e_p' : k_e_p, 'k_e_p_z' : k_e_p_z, 'k_e_p_y' : k_e_p_y, 'k_e_p_d' : k_e_p_d, 'k_carship_p_d' : k_carship_p_d, 
                'k_car_p_z' : k_car_p_z, 'k_car_n_z' : k_car_n_z, 'k_ship_p_z' : k_ship_p_z, 'k_ship_n_z' : k_ship_n_z, 
                'b_ss_n' : b_ss_n, 'u_ss_n' : u_ss_n, 'k_ss_n' : k_ss_n}
    for key, value in dict_index.items():
        hwp.PutFieldText(key, value)


    # 이미지 리스트
    이미지_리스트 = ["chart_1", "chart_2", "chart_3", "chart_4", "chart_5", "chart_6",
                    "chart1", "chart2", "chart3", "chart4", "chart5", "chart6", "chart7", "chart8", "chart9", "chart10", "chart11", "chart12",
                    "chart_a", "chart_b", "chart_c", "chart_d", "chart_e", "chart_f",
                    "s_chart_1", "s_chart_2", "s_chart_3", "s_chart_4", "s_chart_5", "s_chart_6",
                    "i_chart1", "i_chart2", "i_chart3",
                    "car_chart1", "car_chart2", "car_chart3"]

    # 이미지 하나씩 한글 필드에 붙여넣기 (크기 조절 가능)
    for 이미지_이름 in 이미지_리스트:
        images_folder_path = os.path.join(txt_dest_path.get(), 'images')
        이미지_파일_경로 = os.path.join(images_folder_path, f"{이미지_이름}.jpg")
        
        # 이미지 크기 조절
        resize_image(이미지_파일_경로, width_mm=80, height_mm=46)
        
        hwp에_이미지_붙여넣기(hwp, 이미지_이름, 이미지_파일_경로)
        

    # 저장할 파일 경로
    save_path = os.path.join(txt_dest_path.get(), f"{당월} 부울경 경제동향 제출.hwpx")

    # 문서 저장
    hwp.SaveAs(save_path, "HWPX")    
    hwp.Quit()

    # 'images' 폴더 삭제
    if os.path.exists(images_folder_path) and os.path.isdir(images_folder_path):
        try:
            shutil.rmtree(images_folder_path)  # 'images' 폴더 및 하위 항목 삭제
            # print(f"'images' 폴더를 삭제했습니다: {images_folder_path}")
        except OSError as e:
            print(f"Error: {images_folder_path}를 삭제하는 중 오류 발생 - {e}")

    msgbox.showwarning("알림", f"'{당월} 부울경 경제동향 제출.hwpx' 파일이 생성되었습니다." + "\n" + 
                       "   지정한 저장폴더에서 확인 및 실행할 수 있습니다.")
    folder_path = os.path.normpath(txt_dest_path.get())
    subprocess.Popen(['explorer', folder_path], shell=True)



selcol = random.choice(['lightblue1', 'lightcyan','lightyellow2','thistle1','lightgoldenrodyellow','lightsteelblue1'])

## main label frame
main_label = Label(root, bg = selcol, \
    relief = "flat", borderwidth = 2, text = "한글보고서 자동작성 프로그램_for BUSAN", \
    font = ("arial", 15, "bold"), padx = 5, pady = 20)
main_label.pack(side='top', anchor="n", fill = "x")

# 사업체 개인조사표 부문 선택 프레임
choice_label = LabelFrame(root, text = " 이용안내")
choice_label.pack(fill="x", padx=10, pady=10, ipady=5)

describe_label = Label(choice_label, relief = "flat", borderwidth = 0, \
    text = "   ① 업데이트한 엑셀파일 업로드" + "\n" + \
        "   ② 한글보고서 저장할 폴더 선택" + "\n" + \
        "   ③ '파일작성'버튼 클릭" + "\n" + \
        "   ※ 한글파일 최초 자동실행 시 '접근허용' 또는 '모두허용' 클릭 필요!  ",
    font = ("arial", 10, "normal"), padx = 5, pady = 3, justify="left") #justify="left" 구문 왼쪽 정렬
describe_label.pack(side='left', anchor="n", fill = "x")


# 엑셀분석 파일추가 삭제 테두리 프레임 (파일 추가, 선택 삭제)
adel_label = LabelFrame(root, text = " ① 파일추가", bd=1)
adel_label.pack(fill="x", padx=10, pady=10, ipady=5)

# 엑셀분석파일 추가 프레임 (파일 추가, 선택 삭제)
file_frame = Frame(adel_label)
file_frame.pack(fill="x", padx=10, pady=7) # 간격 띄우기

sa_btn_add_file = Button(file_frame, padx=10, pady=5, width=20, text="엑셀분석파일 1개 추가" + "\n" +"(사전 정의된 파일임)", command=sa_add_file)
sa_btn_add_file.pack(side="left", pady=0)

sa_btn_del_file = Button(file_frame, padx=10, pady=5, width=12, text="선택삭제", command=sa_del_file)
sa_btn_del_file.pack(side="right", pady=0)

# 엑셀분석파일 추가를 위한 리스트 프레임
sa_list_frame = Frame(adel_label)
sa_list_frame.pack(fill="both", padx=10, pady=5)

scrollbar = Scrollbar(sa_list_frame)
scrollbar.pack(side="right", fill="y")

sa_list_file = Listbox(sa_list_frame, selectmode="extended", width = 15, height=2, yscrollcommand=scrollbar.set)
sa_list_file.pack(side="left", fill="both", expand=True)
scrollbar.config(command=sa_list_file.yview)

# 드래그 앤 드롭 이벤트를 처리할 바인딩 추가
sa_list_file.drop_target_register(DND_FILES)
sa_list_file.dnd_bind('<<Drop>>', dragDrop)

# 파일 저장 경로 프레임
path_frame = LabelFrame(root, text=" ② 한글파일 저장경로")
path_frame.pack(fill="x", padx=15, pady=10, ipady=5)

txt_dest_path = Entry(path_frame)
txt_dest_path.pack(side="left", fill="x", expand=True, padx=15, pady=7, ipady=4) # 높이 변경

btn_dest_path = Button(path_frame, text="폴더 선택", width=10, command=browse_dest_path)
btn_dest_path.pack(side="right", padx=10, pady=7)


# 실행 프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=10, pady=5)

btn_close = Button(frame_run, padx=10, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right", padx=5, pady=3, anchor = "e")

# 내검실행 버튼 프레임
btn_start = Button(frame_run, bg = "linen", padx=10, pady=5, text=" ③ 파일 작성", width=12, command=start)
btn_start.pack(side="right", padx=5, pady=3, anchor = "e")

root.resizable(True, True)
root.mainloop()