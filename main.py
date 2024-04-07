import tkinter as tk
from tkinter import ttk
from pptx import Presentation
from tkinter import messagebox
from pptx.util import Pt  # Import Pt to set font size
from pptx.dml.color import RGBColor

import pandas as pd
import os
from datetime import datetime
import re

rootdir ="/"

root = tk.Tk()
root.title('말씀 ppt 생성기 - made by eastzoo')
root.geometry("510x340")

bible_range_text = tk.StringVar()

#STYLE
basicFont = 10

# 셀렉트 박스 리스트 담는 변수
bible_list_name = []
# 현재 셀렉트 박스에서 선택된 이름 담는 변수
current_select_bible_name =""
# 사용자가 입력한 성경 범위 저장 함수
bible_range = ""

  # 입력받은 문자를 말씀, 장, 절로 분리하는 함수
def extract_numbers(input_string):
    # 첫번째로 입력받은 문자열의 공백을 모두 제거
    input_string = input_string.replace(" ", "")
    print(input_string)
    # 숫자가 나오기 전까지의 문자열 추출하여 말씀에 저장
    말씀 = ''
    장 = []
    절 = []
    

    # 정규표현식을 사용하여 숫자와 숫자 이후에 ":" 뒤에 나오는 숫자 추출
    pattern = re.compile(r'([^\d:]*)(\d+(?:-\d+)?)(?:[:](\d+(?:-\d+)?))?')
    matches = pattern.match(input_string)

    if matches:
        말씀 = matches.group(1).strip()
        nums = matches.group(2).split('-')
        장.extend([int(num) for num in nums])

        if matches.group(3):
            nums = matches.group(3).split('-')
            절.extend([int(num) for num in nums])
    else:
        말씀 = input_string
    return 말씀, 장, 절

# 입력한 말씀 조건에 맞는 파워포인트 생성 함수
def create_ppt():
    prs = Presentation('template.pptx') # 파워포인트를 다루는 함수 변수 선언Load the PowerPoint template
    # 입력된 말씀 범위 데이터 가져오기   
    slide_layout = prs.slides[0]

    df = pd.read_csv(f"./bible/{current_select_bible_name}.csv")
    df['절'] = pd.to_numeric(df['절'])

    print(df.columns)

    # 입력받은 문자를 말씀, 장, 절로 분리하는 함수
    말씀, 장, 절 = extract_numbers(bible_range_text.get())
    
    print(type(말씀))
    print(type(장[0]))
    print(type(절[0]), 절[1])
    
    print(df.info())
    
    condition = (df.색인 == 말씀) & (df.장 == 장[0]) & (df.절 >= 절[0]) & (df.절 <= 절[1])
    filtered_table = df.loc[ condition ,['색인','장','절','내용']]
    print(df.loc[ condition ,['색인', '장','절','내용']])
    print(len(filtered_table))

    # 첫번째 페이지
    for index, row in filtered_table.iterrows():
        for shape in prs.slides[0].shapes:
            print(shape)
            # 첫번째 페이지의 오브젝트중 글자상자만 찾는 조건문 ( true | false )
            if shape.has_text_frame:
                # 글자상자 변수에 저장
                text_frame = shape.text_frame
                # 기존에 들어있는 글자 삭제
                text_frame.clear()
                # 새로운 글자 입력 ( 말씀 제목 )
                p = text_frame.paragraphs[0]
                
                run = p.add_run()
                run.text = f"{row['색인']} {row['장']}장 {row['절']}절"
                
                # 글꼴 크기 및 글꼴 패밀리 설정
                font = run.font
                font.name = '나눔스퀘어 ExtraBold'
                font.size = Pt(31)
                # 글자색 설정
                font.color.rgb = RGBColor(255,255,255)
                
                prs.slides.add_slide(text_frame)
                


                

                
    # 현재 날짜와 시간을 형식화하여 가져옵니다.
    today_datetime = datetime.today().strftime('%Y-%m-%d %H-%M-%S')
    # Modify the file name to today's date
    new_file_name = f"Presentation_{today_datetime}.pptx"
    # Save the modified presentation
    prs.save(new_file_name)
    messagebox.showinfo("Success", "Presentation created successfully!")


def delete_page():
    for frame in main_frame.winfo_children():
        frame.destroy()
        
        
# 셀렉트 박스 데이터 가져오는 함수
def get_bible_select_data():
    global bible_list_name
    # 성경 폴더에 있는 성경종류 리스트 가져오기 ( 개역개정, + a 예정 )
    bible_list = os.listdir('./bible')
    # 엑셀 확장자 삭제후 보관할 리스트 변수
    for file in bible_list:
        if file.count(".") == 1: 
            name = file.split('.')[0]
            bible_list_name.append(name)
        else:
            for k in range(len(file)-1,0,-1):
                if file[k]=='.':
                    bible_list_name.append(file[:k])
                    break
    return bible_list_name
    
    
# home 페이지 띄우는 함수
def home_page():
    #셀렉트 박스 정보 가져 오기
    get_bible_select_data()
    
    global bible_list_name
    global current_select_bible_name
    
    # 현재 선택한 성경의 변수 저장
    current_select_bible_name = bible_list_name[0]
  
    home_frame = tk.Frame(main_frame)
                    
    # start 셀렉트박스 wrapper 
    ## 성경 셀렉트박스 라벨
    bibile_label = tk.Label(home_frame, text='성경', font=('Bold', basicFont))
    bibile_label.place(x=10, y= 10)

    ## 성경 셀렉트박스
    select_bible = ttk.Combobox(home_frame, values=bible_list_name, width=10)
    select_bible.set(bible_list_name[0])
     # 셀렉트 박스에서 현재 선택된 값 변수에 저장 
    def getSelectedItem(arg):
        global current_select_bible_name
        print(select_bible.get())
        # 선택한 항목 테이블 만드는 함수
        current_select_bible_name = select_bible.get()
    
    select_bible.bind("<<ComboboxSelected>>", getSelectedItem)
    
    select_bible.place(x=50, y=10)
    # end 셀렉트박스 wrapper 
    
    
    # start 성경 범위 입력
    ## 성경범위 라벨
    bibile_range = tk.Label(home_frame, text='성경 구절', font=('Bold', basicFont))
    bibile_range.place(x=10, y= 40)
    ## 성경범위 라벨
    bibile_range_textbox = ttk.Entry(home_frame, textvariable=bible_range_text, width=20)
    bibile_range_textbox.place(x=80, y= 40)
    bibile_range_textbox.focus()
    # end 성경 범위 입력
    
    Fact="""예) 창   =  창세기 전체
창1       =  창세기 1장 전체
롬1-3    =  로마서 1장 1절 - 3장 전체
레1-3:9 =  레위기 1장 1절 - 3장 9절
전1:3    =  전도서 1장 3절
스1:3-9 =  에스라 1장 3절 - 1장 9절
    """
# 사1:3-3:9= 이사야 1장 3절 - 3장 9절
    
    T = tk.Text(home_frame, height = 8, width = 52, )
    T.place(x=10, y=120)
    T.insert(tk.END, Fact)
    
    make_ppt_btn = tk.Button(home_frame, text='PPT 만들기', font=('Bold', basicFont), command=create_ppt)
    make_ppt_btn.place(x=10, y=80)

    home_frame.pack(expand=True, fill='both', pady=20)
    
    
    
# prev 페이지 띄우는 함수
def prev_page():
    prev_frame = tk.Frame(main_frame)

    lb = tk.Label(prev_frame, text='prev Page \n\nPage: 2', font=('Bold', 30))
    lb.pack()

    prev_frame.pack(pady=20)
    
# 사이드 메뉴 선택에 따라 다른 페이지 숨김 함수
def hide_indicators():
    home_indicate.config(bg='#DADADA')
    prev_indicate.config(bg='#DADADA')
    
#사이드 메뉴 클릭시 클릭 css 표시 함수
def indicate(lb, page):
    hide_indicators()
    lb.config(bg='#F15642')
    delete_page()
    page()






# 사이드바  options_frame
options_frame = tk.Frame(root, bg='#DADADA')

options_frame.pack(side=tk.LEFT)
options_frame.pack_propagate(False)
options_frame.configure(width=50, height=400)

# 메인 프레임  main_frame
main_frame = tk.Frame(root, highlightthickness=2)

main_frame.pack(side=tk.LEFT)
main_frame.pack_propagate(False)
main_frame.configure(height=400, width=500)


# 홈 버튼 이미지
#start
homeImg = tk.PhotoImage(file="./assets/icons/ppt.png")
home_btn = tk.Button(options_frame, text='Home', font=('Bold', 15), image=homeImg, height=50, width=50, command=lambda: indicate(home_indicate, home_page))
home_btn.place(x=0, y=0)

home_indicate = tk.Label(options_frame, text='')
home_indicate.place(x=0, y=0, width=5, height=55)
#end


# 기록 버튼 이미지
#start
prevImg = tk.PhotoImage(file="./assets/icons/history.png")
prev_btn = tk.Button(options_frame, text='prev', font=('Bold', 15), image=prevImg, height=50, width=50, command=lambda: indicate(prev_indicate, prev_page))
prev_btn.place(x=0, y=55)

prev_indicate = tk.Label(options_frame, text='')
prev_indicate.place(x=0, y=55, width=5, height=55)
#end


# 처음 시작할때 메인 페이지 세팅 __INIT__
indicate(home_indicate, home_page)
root.mainloop() 