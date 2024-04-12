#custom modules
import query
import tkinter as tk
from tkinter import ttk

# prev 페이지 띄우는 함수
def history_page(main_frame):
    history_list = query.select()
    history_frame = tk.Frame(main_frame)
    
    
    def show_selected_item():
        # 선택한 아이템의 ID 가져오기
        selected_item = ListTree.selection()
        if selected_item:
            # 선택한 아이템의 정보 가져오기
            item = ListTree.item(selected_item)
            values = item['values']
            print("Selected item:", values)
        else:
            print("No item selected")
    
    # 리스트 리패칭
    def refresh_list():
        # 현재 트리뷰의 모든 항목 삭제
        ListTree.delete(*ListTree.get_children())
        # 삭제 후 리패칭
        history_list = query.select()
     
        # 데이터를 리스트 박스에 추가
        for item in history_list:
            ListTree.insert('',"end", values=(f'{item[0]}',f'{item[1]}',f'{item[2]}'))

    def delete_selected_item():
        selected_item = ListTree.selection()
        if selected_item:
            # 선택한 아이템의 정보 가져오기
            item = ListTree.item(selected_item)
            values = item['values']
            # id를 uniq 키로 전송하여 삭제
            query.delete(values[0])
        else:
            print("No item selected")
        #삭제 리패칭
        refresh_list()
    
    # 리스트 박스 생성
    # padx = 좌우 패딩
    ListTree = ttk.Treeview(history_frame, column=("c1", "c2", "c3"), show='headings', height=5)
    ListTree.pack(side="left", fill="y",padx=10)
    
    ListTree.column("# 1", anchor="center", width=30)
    ListTree.heading("# 1", text="id")
    ListTree.column("# 2", anchor="center",width=150)
    ListTree.heading("# 2", text="내용")
    ListTree.column("# 3", anchor="center",width=150)
    ListTree.heading("# 3", text="날짜")

    # 데이터를 리스트 박스에 추가
    for item in history_list:
        ListTree.insert('',"end", values=(f'{item[0]}',f'{item[1]}',f'{item[2]}'))

    # 확인 버튼 생성
    button = tk.Button(history_frame, text="불러오기", command=show_selected_item)
    button.pack(side="top", fill="x")
    
     # 삭제 버튼 생성
    button = tk.Button(history_frame, text="삭제", command=delete_selected_item)
    button.pack(side="top", fill="x")

    history_frame.pack(expand=True, fill='both', pady=20)

    