#custom modules
import query
import tkinter as tk


# prev 페이지 띄우는 함수
def history_page(main_frame):
    history_list = query.select()
    history_frame = tk.Frame(main_frame)
    
    
    def show_selected_item():
        selected_item = listbox.get(listbox.curselection())
        print("Selected item:", selected_item)
    
    

    print("history_list", history_list)
    # 스크롤 가능한 리스트 박스 생성
    listbox = tk.Listbox(history_frame, height=5, selectmode="single")
    listbox.pack(side="left", fill="y")

    scrollbar = tk.Scrollbar(history_frame, orient="vertical", command=listbox.yview)
    scrollbar.pack(side="right", fill="y")
    listbox.config(yscrollcommand=scrollbar.set)

    # 데이터를 리스트 박스에 추가
    for item in history_list:
        listbox.insert("end", f"{item[1]}, {item[2]}")

    # 확인 버튼 생성
    button = tk.Button(history_frame, text="확인", command=show_selected_item)
    button.pack()

    history_frame.pack(expand=True, fill='both', pady=20)

    