
#=====================================
import tkinter as tk
import tkinter.ttk as ttk
import webbrowser
import pandas as pd
import json
from tkinter import filedialog
import webbrowser

# 데이터 파일 경로 (변경 가능)
DATA_FILE = "vessel_data.json"

# 초기 데이터 로드 (파일에서 로드하거나 빈 리스트로 초기화)
try:
    with open(DATA_FILE, 'r') as f:
        data = json.load(f)
except FileNotFoundError:
    data = []


def exit_program():
    save_data()  # 데이터 저장 후 종료
    window.destroy()

# def open_websites():
#     keywords = entry.get("1.0", tk.END).splitlines()  # 텍스트 상자 내용을 줄바꿈으로 분리

def open_websites():
    keywords = entry.get("1.0", tk.END).splitlines()  # 텍스트 상자 내용을 줄바꿈으로 분리
    
    for keyword in keywords:
        keyword = keyword.strip()  # 띄어쓰기 제거
        
        for item in data:
            if keyword.lower() in item[0].lower():  # 대소문자 구분 없이 비교
                webbrowser.open_new_tab(item[1])
                break  # 일치하는 항목 찾으면 다음 키워드로 이동
        else:
            # 일치하는 키워드가 없을 경우 처리 (예: 메시지 출력)
            print(f"키워드 '{keyword}'와 일치하는 주소가 없습니다.")

def open_websites2():

    keywords = entry2.get("1.0", tk.END).splitlines()  # 텍스트 상자 내용을 줄바꿈으로 분리

    search_vsl_name = []
    for vsl in keywords:
        vsl = vsl.strip()  # 띄어쓰기 제거
        last_space_index = vsl.rfind(' ')
        processed_line = vsl[:last_space_index] if last_space_index != -1 else vsl
        search_vsl_name.append(processed_line)

    for keyword in search_vsl_name:
        keyword = keyword.strip()  # 띄어쓰기 제거
        
        for item in data:
            if keyword.lower() in item[0].lower():  # 대소문자 구분 없이 비교
                webbrowser.open_new_tab(item[1])
                break  # 일치하는 항목 찾으면 다음 키워드로 이동
        else:
            # 일치하는 키워드가 없을 경우 처리 (예: 메시지 출력)
            print(f"키워드 '{keyword}'와 일치하는 주소가 없습니다.")

def import_excel_data():
    global data
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path, header=None, names=["keyword", "url"])
            new_data = df.values.tolist()
            data.extend(new_data) 
            update_treeview()
            data_count_label.config(text=f"Total data: {len(data)}")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Excel 파일을 읽는 동안 오류가 발생했습니다: {e}")

def update_treeview():
    tree.delete(*tree.get_children()) 
    for item in data:
        tree.insert("", tk.END, values=item)

def save_data():
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f)

def open_bl_websites():
    bl_numbers = bl_entry.get("1.0", tk.END).splitlines()
    container_count = 0
    bl_count = 0
    for bl_number in bl_numbers:
        bl_number = bl_number.strip()
        if len(bl_number) == 11:
            container_count += 1
        elif len(bl_number) == 12:
            bl_count += 1
        else:
            continue  # 유효하지 않은 길이의 번호는 무시
        
        url = "https://teamsite.msc.com/sites/together/queries/pages/home.aspx#!/logistics/tracker-history?searchBy=CTBL&searchVal=" + bl_number
        webbrowser.open_new_tab(url)

def count_containers_and_bls(event):
    text = bl_entry.get("1.0", tk.END)  # 전체 텍스트 가져오기
    lines = text.splitlines()  # 줄바꿈으로 분리
    container_count = 0
    bl_count = 0

    for line in lines:
        line = line.strip()  # 각 줄에서 공백 제거
        if len(line) == 11:
            container_count += 1
        elif len(line) == 12:
            bl_count += 1

    count_label.config(text=f"Container: {container_count}, BL: {bl_count}")

# 웹사이트 열기 함수 (BL Search)        
def open_bl_websites():
    bl_numbers = bl_entry.get("1.0", tk.END).splitlines()
    for bl_number in bl_numbers:
        bl_number = bl_number.strip()
        url = "https://teamsite.msc.com/sites/together/queries/pages/home.aspx#!/logistics/tracker-history?searchBy=CTBL&searchVal=" + bl_number
        webbrowser.open_new_tab(url)

def search_and_filter2(treeview, search_entry):
    """트리뷰에서 검색어로 필터링된 항목을 강조합니다."""
    search_terms = search_entry.get("1.0", tk.END).strip().upper().splitlines()  # 대소문자 구분 없이 검색어 분리
    all_ids = treeview.get_children()

    search_vsl_name = []
    for vsl in search_terms:
        last_space_index = vsl.rfind(' ')
        processed_line = vsl[:last_space_index] if last_space_index != -1 else vsl
        search_vsl_name.append(processed_line)

    search_vsl_voy = []
    for vsl in search_terms:
        processed_line = vsl.split()[-1]
        search_vsl_voy.append(processed_line)

    # 모든 항목 태그 초기화
    for child_id in all_ids:
        treeview.item(child_id, tags=())  # 기존 태그 제거

    # 일치하는 항목 강조 표시
    found_items = 0
    for child_id in all_ids:
        values = treeview.item(child_id)['values']
        if any(term in str(value).upper() for term in search_vsl_name for value in values):
            for voy in search_vsl_voy:
                if voy in values[3]:
                    treeview.item(child_id, tags=('found',))
                    found_items += 1

    if found_items > 0:
        treeview.tag_configure('found', background='greenyellow')


# Tkinter 창 생성
window = tk.Tk()
window.title("Vessel/BL/Container Search")
window.geometry("500x400")  # 창 크기 설정
 
style = ttk.Style()

# 탭 생성
notebook = ttk.Notebook(window)
notebook.pack(fill='both', expand=True)

# 검색 탭 생성
search_tab = ttk.Frame(notebook)
notebook.add(search_tab, text="Vessel Search")



# 입력 상자 생성

entry = tk.Text(search_tab)
entry.pack()
entry.place(x=25, y=80, height=300, width=130) 

data_listbox = tk.Listbox(search_tab, width=50,background='khaki') 
data_listbox.pack(side=tk.RIGHT, fill=tk.Y)  # Fill vertically only
data_listbox.place(x=180, y=80, height=300, width=130) 

entry2 = tk.Text(search_tab)
entry2.pack()
entry2.place(x=335, y=80, height=300, width=130) 


# 버튼 생성
button = tk.Button(search_tab, text="Vessel only", command=open_websites)
button.pack()
button.place(x=50, y=43, height=30, width=80)

# 버튼 생성2
button2 = tk.Button(search_tab, text="Vessel + Voy", command=open_websites2)
button2.pack()
button2.place(x=365, y=43, height=30, width=80) 

# 데이터 탭 생성
data_tab = ttk.Frame(notebook)
notebook.add(data_tab, text="Vessel Data")

# Treeview 생성
tree = ttk.Treeview(data_tab, columns=("Keyword", "URL"), show="headings")
tree.heading("Keyword", text="Keyword")
tree.heading("URL", text="URL")
tree.column("Keyword", width=200)  # Keyword 열 폭 설정
tree.pack()

# 초기 데이터 표시
update_treeview()

# 데이터 갯수 레이블
data_count_label = tk.Label(data_tab, text=f"Total data: {len(data)}")
data_count_label.pack()

# Excel 가져오기 버튼
import_button = tk.Button(data_tab, text="Import Excel Data", command=import_excel_data)
import_button.pack()

# 저장 버튼
save_button = tk.Button(data_tab, text="Save Data", command=save_data)
save_button.pack()

exit_button = tk.Button(search_tab, text="종료", command=exit_program)
exit_button.pack()



def update_listbox():
    data_listbox.delete(0, tk.END)  # Clear the listbox
    for item in data:
        data_listbox.insert(tk.END, f"{item[0]}")  # Add items to the listbox

update_listbox()



# Container/BL Search 탭 생성
bl_search_tab = ttk.Frame(notebook)
notebook.add(bl_search_tab, text="Container/BL Search")


# 입력 내용이 변경될 때마다 길이 확인


# 입력 상자 생성 (BL Search 탭)
bl_label = tk.Label(bl_search_tab, text="Container/BL 번호 입력 (한 줄에 하나씩):")
bl_label.pack()

bl_entry = tk.Text(bl_search_tab,background='pink')
bl_entry.pack()
bl_entry.place(x=185, y=80, height=300, width=130)

# 문자 수를 표시할 레이블 생성
# count_label = tk.Label(bl_search_tab, text="SADFSADFSADF")
# count_label.pack()
# count_label.place(x=350, y=80, height=20, width=120) 

count_label = tk.Label(bl_search_tab, text="")
count_label.pack()
count_label.place(x=350, y=80, height=20, width=120) 

bl_entry.bind("<KeyRelease>", count_containers_and_bls)  # 키 입력 시 함수 호출


# 버튼 생성 (BL Search)
bl_button = tk.Button(bl_search_tab, text="Container/BL", command=open_bl_websites)
bl_button.pack()

exit_button = tk.Button(bl_search_tab, text="종료", command=exit_program)
exit_button.pack()


# 프로그램 종료 시 데이터 저장
window.protocol("WM_DELETE_WINDOW", save_data)

window.mainloop()