import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import json
from typing import List, Dict
from tkinterdnd2 import DND_FILES, TkinterDnD
from datetime import datetime
from openpyxl import load_workbook
import xlrd  # .xls 파일 처리를 위한 라이브러리

# pyinstaller -w -F --add-binary="C:/Users/kod03/AppData/Local/Programs/Python/Python311/tcl/tkdnd2.8;tkdnd2.8" obl_project1.py

class ContainerConverter:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("CLL to OBL Converter")
        self.root.geometry("1000x900")  # 창 크기 증가

        # 설정 파일 경로 설정
        user_home = os.environ['USERPROFILE']  # Windows 사용자 프로필 경로
        self.config_dir = os.path.normpath(os.path.join(user_home, "Desktop", "OBL_Configs"))
        os.makedirs(self.config_dir, exist_ok=True)
        
        self.stowage_config_file = os.path.normpath(os.path.join(self.config_dir, "StowCodes_mapping.json"))  # 파일명 수정
        self.port_config_file = os.path.normpath(os.path.join(self.config_dir, "port_mapping.json"))
        self.tpsz_config_file = os.path.normpath(os.path.join(self.config_dir, "SZTP_mapping.json"))  # TPSZ 파일명도 수정
        
        # 매핑 설정 로드
        self.stowage_settings = self.load_stowage_settings()
        self.stow_mapping = self.stowage_settings.get('mapping', {})
        self.stow_column_mapping = self.stowage_settings.get('column_mapping', {
            'discharge_port': 'FPOD',
            'stowage_code': 'Stow'
        })
        
        # TpSz 매핑 설정 로드
        self.tpsz_settings = self.load_tpsz_settings()
        self.tpsz_mapping = self.tpsz_settings.get('mapping', {})
        self.tpsz_column_mapping = self.tpsz_settings.get('column_mapping', {
            'before': 'Description',
            'after': 'Code'
        })
        
        # PORT CODE 매핑 추가
        self.port_codes = {'AEAJM': 'AJMAN'}
        self.current_file = None
        self.output_file = None

        # POL, TOL 선택 값 저장 변수
        self.selected_pol = tk.StringVar()
        self.selected_tol = tk.StringVar()

        # ITPS 관련 변수
        self.itps_file = None
        self.obl_file = None

        self.setup_ui()
        self.reset_all()  # 프로그램 시작 시 자동으로 초기화 실행

    def load_stowage_settings(self) -> Dict:
        try:
            # 하드코딩된 바탕화면 경로 사용
            config_file = os.path.normpath(r"C:\\Users\\kod03\\OneDrive\\바탕 화면\\StowCodes_mapping.json")
            with open(config_file, 'r') as f:
                return json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Stowage 매핑 파일 로드 실패: {str(e)}")
            return {}

    def load_tpsz_settings(self) -> Dict:
        try:
            # 하드코딩된 바탕화면 경로 사용  
            config_file = os.path.normpath(r"C:\\Users\\kod03\\OneDrive\\바탕 화면\\SZTP_mapping.json")
            with open(config_file, 'r') as f:
                return json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"TpSz 매핑 파일 로드 실패: {str(e)}")
            return {}
    #===================================================================================================
    # def load_stowage_settings(self) -> Dict:
    #     try:
    #         with open(self.stowage_config_file, 'r') as f:
    #             return json.load(f)
    #     except Exception as e:
   
    #         messagebox.showerror("Error", f"Stowage 매핑 파일 로드 실패: {str(e)}")
    #         return {}

    # def load_tpsz_settings(self) -> Dict:
    #     try:
    #         with open(self.tpsz_config_file, 'r') as f:
    #             return json.load(f)
    #     except Exception as e:
    #         messagebox.showerror("Error", f"TpSz 매핑 파일 로드 실패: {str(e)}")
    #         return {}
    #===================================================================================================


    def setup_ui(self):
        # 탭 컨트롤 생성
        self.tab_control = ttk.Notebook(self.root)
        self.tab_control.pack(expand=True, fill="both")
        
        # 초기화 버튼 추가 (오른쪽 상단)
        reset_frame = ttk.Frame(self.root)
        reset_frame.pack(anchor='ne', padx=10, pady=5)
        
        reset_button = ttk.Button(reset_frame, text="초기화", command=self.reset_all)
        reset_button.pack()
        
        # 기존 단일 CLL 변환 탭
        self.single_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.single_tab, text='단일 CLL 변환')
        
        # Multi CLL 변환 탭
        self.multi_cll_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.multi_cll_tab, text='Multi CLL 변환')
        
        # ITPS 추가 탭
        self.itps_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.itps_tab, text='ITPS 추가')

        # STOWAGE CODE 관리 탭
        self.stowage_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.stowage_tab, text='STOWAGE CODE 관리')

        # TpSZ 관리 탭
        self.tpsz_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tpsz_tab, text='TpSZ 관리')

        # 각 탭 설정
        self.setup_single_tab()  # 단일 CLL 탭 설정
        self.setup_multi_cll_tab()  # Multi CLL 탭 설정
        self.setup_itps_tab()  # ITPS 탭 설정
        self.setup_stowage_tab()  # STOWAGE CODE 탭 설정
        self.setup_tpsz_tab()  # TpSZ 탭 설정

    def setup_single_tab(self):
        # 단일 CLL 변환 탭 설정
        left_frame = ttk.Frame(self.single_tab)
        left_frame.pack(side="left", fill="both", expand=True, padx=5)
        
        right_frame = ttk.Frame(self.single_tab)
        right_frame.pack(side="right", fill="both", padx=5)
        
        # POL, TOL 선택 프레임
        port_frame = ttk.LabelFrame(left_frame, text="POL TOL")
        port_frame.pack(pady=10, padx=10, fill="x")

        # POL 버튼 프레임
        pol_frame = ttk.LabelFrame(port_frame, text="POL")
        pol_frame.pack(pady=5, padx=5, fill="x")

        # POL 버튼들
        pol_ports = ['KRPUS', 'KRKAN', 'KRINC']
        self.pol_buttons = {}
        for port in pol_ports:
            btn = tk.Button(pol_frame, text=port, width=10,
                          command=lambda p=port: self.select_pol(p))
            btn.pack(side=tk.LEFT, padx=5, pady=5)
            self.pol_buttons[port] = btn

        # TOL 버튼 프레임
        tol_frame = ttk.LabelFrame(port_frame, text="TOL")
        tol_frame.pack(pady=5, padx=5, fill="x")

        # TOL 버튼들과 매핑
        tol_mapping = {
            'PNC': 'KRPUSPN',
            'PNIT': 'KRPUSAB',
            'BCT': 'KRPUSBC',
            'HJNC': 'KRPUSAP',
            'ICT': 'KRINCAH',
            'GWCT': 'KRKANKT'
        }
        self.tol_buttons = {}
        self.tol_values = tol_mapping
        for btn_text, value in tol_mapping.items():
            btn = tk.Button(tol_frame, text=btn_text, width=10,
                          command=lambda v=value: self.select_tol(v))
            btn.pack(side=tk.LEFT, padx=5, pady=5)
            self.tol_buttons[value] = btn

        # 파일 정보 표시 영역
        info_frame = ttk.LabelFrame(left_frame, text="파일 정보")
        info_frame.pack(pady=10, padx=10, fill="x")

        self.input_label = ttk.Label(info_frame, text="입력 파일: 없음")
        self.input_label.pack(pady=5, anchor="w")

        self.output_label = ttk.Label(info_frame, text="출력 파일: 없음")
        self.output_label.pack(pady=5, anchor="w")

        # CLL 변환을 위한 드래그 & 드롭 영역
        self.cll_frame = ttk.LabelFrame(left_frame, text="CLL -> OBL 변환")
        self.cll_frame.pack(pady=10, padx=10, fill="x")

        self.cll_label = ttk.Label(self.cll_frame, text="CLL 파일을 여기에 드롭하세요")
        self.cll_label.pack(pady=20)

        # CLL 드래그 앤 드롭 바인딩
        self.cll_label.drop_target_register(DND_FILES)
        self.cll_label.dnd_bind('<<Drop>>', self.drop_cll_file)

        # OBL EMPTY 추가를 위한 드래그 & 드롭 영역
        self.obl_frame = ttk.LabelFrame(left_frame, text="OBL EMPTY 추가")
        self.obl_frame.pack(pady=10, padx=10, fill="x")

        self.obl_label = ttk.Label(self.obl_frame, text="OBL 파일을 여기에 드롭하세요")
        self.obl_label.pack(pady=20)

        # OBL 드래그 앤 드롭 바인딩
        self.obl_label.drop_target_register(DND_FILES)
        self.obl_label.dnd_bind('<<Drop>>', self.drop_obl_file)

        # EMPTY 컨테이너 입력 섹션
        empty_frame = ttk.LabelFrame(left_frame, text="EMPTY 컨테이너 추가")
        empty_frame.pack(pady=10, padx=10, fill="x")

        # 5개의 입력 행 생성
        self.empty_entries = []
        for i in range(5):
            row_frame = ttk.Frame(empty_frame)
            row_frame.pack(pady=5)

            pod_entry = ttk.Entry(row_frame, width=10)
            pod_entry.pack(side="left", padx=5)
            pod_entry.insert(0, "POD")
            pod_entry.bind('<FocusIn>', lambda e, entry=pod_entry: self.on_entry_click(e, entry))
            pod_entry.bind('<FocusOut>', lambda e, entry=pod_entry: self.on_focus_out(e, entry, "POD"))
            pod_entry.bind('<Key>', lambda e, entry=pod_entry: self.on_key_press(e, entry))

            sztp_entry = ttk.Entry(row_frame, width=10)
            sztp_entry.pack(side="left", padx=5)
            sztp_entry.insert(0, "SzTp")
            sztp_entry.bind('<FocusIn>', lambda e, entry=sztp_entry: self.on_entry_click(e, entry))
            sztp_entry.bind('<FocusOut>', lambda e, entry=sztp_entry: self.on_focus_out(e, entry, "SzTp"))
            sztp_entry.bind('<Key>', lambda e, entry=sztp_entry: self.on_key_press(e, entry))

            qty_entry = ttk.Entry(row_frame, width=5)
            qty_entry.pack(side="left", padx=5)
            qty_entry.insert(0, "수량")
            qty_entry.bind('<FocusIn>', lambda e, entry=qty_entry: self.on_entry_click(e, entry))
            qty_entry.bind('<FocusOut>', lambda e, entry=qty_entry: self.on_focus_out(e, entry, "수량"))
            qty_entry.bind('<Key>', lambda e, entry=qty_entry: self.on_key_press(e, entry))

            self.empty_entries.append((pod_entry, sztp_entry, qty_entry))

        # Summary 표시 영역을 right_frame으로 이동
        self.single_summary_frame = ttk.LabelFrame(right_frame, text="Container Summary")
        self.single_summary_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.single_summary_text = tk.Text(self.single_summary_frame, height=30, width=40)
        self.single_summary_text.pack(pady=5, padx=5, fill="both", expand=True)
        self.single_summary_text.insert(tk.END, "단일 CLL 탭에서 파일 변환 시 Summary가 표시됩니다.")

    def setup_multi_cll_tab(self):
        """CLL 파일 병합 탭 설정"""
        # 좌우 분할
        left_frame = ttk.Frame(self.multi_cll_tab)
        left_frame.pack(side="left", fill="both", expand=True, padx=5)
        
        right_frame = ttk.Frame(self.multi_cll_tab)
        right_frame.pack(side="right", fill="both", padx=5)
        
        # POL/TOL 선택 프레임
        port_frame = ttk.LabelFrame(left_frame, text="POL TOL")
        port_frame.pack(pady=10, padx=10, fill="x")
        
        # POL 버튼 프레임
        pol_frame = ttk.LabelFrame(port_frame, text="POL")
        pol_frame.pack(pady=5, padx=5, fill="x")
        
        pol_ports = ['KRPUS', 'KRKAN', 'KRINC']
        self.multi_pol_buttons = {}
        for port in pol_ports:
            btn = tk.Button(pol_frame, text=port, width=10,
                          command=lambda p=port: self.select_multi_pol(p))
            btn.pack(side=tk.LEFT, padx=5, pady=5)
            self.multi_pol_buttons[port] = btn

        # TOL 버튼 프레임
        tol_frame = ttk.LabelFrame(port_frame, text="TOL")
        tol_frame.pack(pady=5, padx=5, fill="x")
        
        tol_mapping = {
            'PNC': 'KRPUSPN',
            'PNIT': 'KRPUSAB',
            'BCT': 'KRPUSBC',
            'HJNC': 'KRPUSAP',
            'ICT': 'KRINCAH',
            'GWCT': 'KRKANKT'
        }
        
        self.multi_tol_buttons = {}
        for btn_text, value in tol_mapping.items():
            btn = tk.Button(tol_frame, text=btn_text, width=10,
                          command=lambda v=value: self.select_multi_tol(v))
            btn.pack(side=tk.LEFT, padx=5, pady=5)
            self.multi_tol_buttons[btn_text] = btn

        # 파일 선택 영역 컨테이너
        files_frame = ttk.Frame(left_frame)
        files_frame.pack(pady=10, padx=10, fill="x")

        # Master CLL 파일 프레임
        self.master_frame = ttk.LabelFrame(files_frame, text="첫 번째(Master) CLL 파일")
        self.master_frame.pack(pady=5, padx=5, fill="x")
        
        self.master_label = ttk.Label(self.master_frame, text="CLL 파일을 여기에 드롭하세요")
        self.master_label.pack(pady=10)
        
        self.master_path_label = ttk.Label(self.master_frame, text="파일 경로: 없음")
        self.master_path_label.pack(pady=5)
        
        # Master 파일 드롭 영역 바인딩
        self.master_frame.drop_target_register(DND_FILES)
        self.master_frame.dnd_bind('<<Drop>>', self.drop_master_cll)

        # Slave CLL 파일 프레임
        self.slave_frame = ttk.LabelFrame(files_frame, text="두 번째(Slave) CLL 파일")
        self.slave_frame.pack(pady=5, padx=5, fill="x")
        
        self.slave_label = ttk.Label(self.slave_frame, text="CLL 파일을 여기에 드롭하세요")
        self.slave_label.pack(pady=10)
        
        self.slave_path_label = ttk.Label(self.slave_frame, text="파일 경로: 없음")
        self.slave_path_label.pack(pady=5)
        
        # Slave 파일 드롭 영역 바인딩
        self.slave_frame.drop_target_register(DND_FILES)
        self.slave_frame.dnd_bind('<<Drop>>', self.drop_slave_cll)

        # 결과 정보 프레임
        self.result_frame = ttk.LabelFrame(right_frame, text="변환 결과")
        self.result_frame.pack(pady=10, padx=10, fill="x")
        
        self.result_label = ttk.Label(self.result_frame, text="출력 파일: 없음")
        self.result_label.pack(pady=5)

        # Summary 표시 영역을 right_frame으로 이동
        self.multi_summary_frame = ttk.LabelFrame(right_frame, text="Container Summary")
        self.multi_summary_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.multi_summary_text = tk.Text(self.multi_summary_frame, height=30, width=40)
        self.multi_summary_text.pack(pady=5, padx=5, fill="both", expand=True)
        self.multi_summary_text.insert(tk.END, "Multi CLL 탭에서 파일 변환 시 Summary가 표시됩니다.")

    def setup_itps_tab(self):
        """ITPS 추가 탭 설정"""
        # 좌우 분할
        left_frame = ttk.Frame(self.itps_tab)
        left_frame.pack(side="left", fill="both", expand=True, padx=5)
        
        right_frame = ttk.Frame(self.itps_tab)
        right_frame.pack(side="right", fill="both", padx=5)

        # 파일 정보 표시 영역
        info_frame = ttk.LabelFrame(left_frame, text="파일 정보")
        info_frame.pack(pady=10, padx=10, fill="x")

        self.itps_input_label = ttk.Label(info_frame, text="ITPS 파일: 없음")
        self.itps_input_label.pack(pady=5, anchor="w")

        self.itps_obl_label = ttk.Label(info_frame, text="OBL 파일: 없음")
        self.itps_obl_label.pack(pady=5, anchor="w")

        self.itps_output_label = ttk.Label(info_frame, text="출력 파일: 없음")
        self.itps_output_label.pack(pady=5, anchor="w")

        # ITPS 파일 드롭 영역
        itps_drop_frame = ttk.LabelFrame(left_frame, text="ITPS 파일 드롭")
        itps_drop_frame.pack(pady=10, padx=10, fill="x")

        self.itps_drop_label = ttk.Label(itps_drop_frame, text="ITPS 파일을 여기에 드롭하세요")
        self.itps_drop_label.pack(pady=20)

        # ITPS 드래그 앤 드롭 바인딩
        self.itps_drop_label.drop_target_register(DND_FILES)
        self.itps_drop_label.dnd_bind('<<Drop>>', self.drop_itps_file)

        # OBL 파일 드롭 영역
        obl_drop_frame = ttk.LabelFrame(left_frame, text="OBL 파일 드롭")
        obl_drop_frame.pack(pady=10, padx=10, fill="x")

        self.obl_drop_label = ttk.Label(obl_drop_frame, text="OBL 파일을 여기에 드롭하세요")
        self.obl_drop_label.pack(pady=20)

        # OBL 드래그 앤 드롭 바인딩
        self.obl_drop_label.drop_target_register(DND_FILES)
        self.obl_drop_label.dnd_bind('<<Drop>>', self.drop_obl_for_itps)

        # Summary 표시 영역
        self.itps_summary_frame = ttk.LabelFrame(right_frame, text="ITPS Summary")
        self.itps_summary_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.itps_summary_text = tk.Text(self.itps_summary_frame, height=30, width=40)
        self.itps_summary_text.pack(pady=5, padx=5, fill="both", expand=True)
        self.itps_summary_text.insert(tk.END, "ITPS 파일 처리 시 Summary가 표시됩니다.")

    def setup_stowage_tab(self):
        """STOWAGE CODE 관리 탭 설정"""
        # 메인 프레임
        main_frame = ttk.Frame(self.stowage_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 서비스 선택 프레임
        service_frame = ttk.LabelFrame(main_frame, text="Service Name 선택")
        service_frame.pack(fill="x", pady=(0, 10))
        
        # 서비스 선택 콤보박스
        self.selected_service = tk.StringVar()
        self.service_combo = ttk.Combobox(service_frame, textvariable=self.selected_service)
        self.service_combo.pack(pady=10, padx=5, fill="x")
        self.service_combo.bind('<<ComboboxSelected>>', self.on_service_selected)

        # 드래그 & 드롭 영역
        drop_frame = ttk.LabelFrame(main_frame, text="Stowage Code 엑셀 파일")
        drop_frame.pack(fill="x", pady=(0, 10))

        self.stowage_drop_label = ttk.Label(drop_frame, text="Stowage Code 엑셀 파일을 여기에 드롭하세요")
        self.stowage_drop_label.pack(pady=20)

        # 드래그 앤 드롭 바인딩
        self.stowage_drop_label.drop_target_register(DND_FILES)
        self.stowage_drop_label.dnd_bind('<<Drop>>', self.drop_stowage_file)

        # 컬럼 매핑 설정 영역
        mapping_frame = ttk.LabelFrame(main_frame, text="컬럼 매핑 설정")
        mapping_frame.pack(fill="x", pady=(0, 10))

        # Discharge Port 컬럼 매핑
        discharge_frame = ttk.Frame(mapping_frame)
        discharge_frame.pack(fill="x", pady=5)
        ttk.Label(discharge_frame, text="Discharge Port 컬럼명:").pack(side="left", padx=5)
        self.discharge_entry = ttk.Entry(discharge_frame)
        self.discharge_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.discharge_entry.insert(0, self.stow_column_mapping.get('discharge_port', ''))

        # Stowage Code 컬럼 매핑
        stowage_frame = ttk.Frame(mapping_frame)
        stowage_frame.pack(fill="x", pady=5)
        ttk.Label(stowage_frame, text="Stowage Code 컬럼명:").pack(side="left", padx=5)
        self.stowage_entry = ttk.Entry(stowage_frame)
        self.stowage_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.stowage_entry.insert(0, self.stow_column_mapping.get('stowage_code', ''))

        # 저장 버튼
        save_button = ttk.Button(mapping_frame, text="설정 저장", command=self.save_stowage_settings)
        save_button.pack(pady=10)

        # 현재 매핑 미리보기
        preview_frame = ttk.LabelFrame(main_frame, text="현재 매핑 미리보기")
        preview_frame.pack(fill="both", expand=True)
        
        self.preview_text = tk.Text(preview_frame, height=10)
        self.preview_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 현재 매핑 표시
        self.update_stowage_preview()

    def on_service_selected(self, event):
        """서비스 선택 시 처리"""
        self.update_stowage_preview()

    def on_entry_click(self, event, entry):
        """Entry 위젯 클릭시 기본 텍스트 제거"""
        if entry.get() in ["POD", "SzTp", "수량"]:
            entry.delete(0, tk.END)
            entry.config(foreground='black')

    def on_focus_out(self, event, entry, default_text):
        """Entry 위젯에서 포커스가 빠졌을 때 처리"""
        if entry.get().strip() == "":
            entry.insert(0, default_text)
            entry.config(foreground='gray')

    def on_key_press(self, event, entry):
        """키 입력 처리"""
        if entry.get() in ["POD", "SzTp", "수량"]:
            entry.delete(0, tk.END)

    def on_tab(self, event):
        """탭 키 처리"""
        current = event.widget
        next_widget = current.tk_focusNext()
        next_widget.focus()
        return "break"  # 기본 탭 동작 방지

    def find_matching_services(self, pod_list):
        """POD 리스트와 매칭되는 서비스 찾기"""
        matching_services = {}
        for service_name, mappings in self.stow_mapping.items():
            for pod in pod_list:
                for mapping in mappings:
                    if pod.upper() == mapping['port'].upper() or pod.upper() == mapping['stow_code'].upper():
                        if service_name not in matching_services:
                            matching_services[service_name] = set()
                        matching_services[service_name].add(pod)
        return matching_services

    def show_service_selection_dialog(self, matching_services):
        """서비스 선택 다이얼로그 표시"""
        dialog = tk.Toplevel(self.root)
        dialog.title("서비스 선택")
        dialog.geometry("400x300")
        
        # 설명 레이블
        ttk.Label(dialog, text="발견된 POD와 일치하는 서비스입니다.\n사용할 서비스를 선택해주세요:", 
                  justify=tk.CENTER).pack(pady=10)
        
        # 서비스 목록 프레임
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 서비스 목록 표시
        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # 서비스와 매칭된 POD 정보 추가
        for service, pods in matching_services.items():
            listbox.insert(tk.END, f"{service} (매칭 POD: {', '.join(pods)})")
        
        # 선택 결과를 저장할 변수
        self.selected_service = tk.StringVar()
        
        def on_select():
            selection = listbox.curselection()
            if selection:
                # 선택된 서비스 이름만 추출 (POD 정보 제외)
                service_name = listbox.get(selection[0]).split(" (")[0]
                self.selected_service.set(service_name)
                dialog.destroy()
        
        # 선택 버튼
        ttk.Button(dialog, text="선택", command=on_select).pack(pady=10)
        
        # 다이얼로그가 닫힐 때까지 대기
        dialog.transient(self.root)
        dialog.grab_set()
        self.root.wait_window(dialog)
        
        return self.selected_service.get()

    def drop_cll_file(self, event):
        """단일 CLL 파일 드롭 처리"""
        file_path = event.data.strip('{}').strip('"')
        if not os.path.exists(file_path):
            messagebox.showerror("오류", "파일이 존재하지 않습니다.")
            return

        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(file_path, header=4)
            
            # POD 목록 추출
            pod_list = df['POD'].unique().tolist()
            
            # 매칭되는 서비스 찾기
            matching_services = self.find_matching_services(pod_list)
            
            if not matching_services:
                messagebox.showwarning("경고", "POD와 일치하는 서비스를 찾을 수 없습니다.")
                return
                
            # 서비스 선택 다이얼로그 표시
            selected_service = self.show_service_selection_dialog(matching_services)
            
            if not selected_service:
                return
                
            # 터미널 코드 읽기
            df_check = pd.read_excel(file_path, header=None)
            terminal_code = str(df_check.iloc[3, 11]).strip()

            if not terminal_code:
                messagebox.showerror("오류", "(4,12) 위치에서 터미널 코드를 찾을 수 없습니다.")
                return

            # 터미널 코드를 기반으로 POL, TOL 값 자동 설정
            port_info = self.terminal_to_port_mapping(terminal_code)
            
            if not port_info['pol'] or not port_info['tol']:
                messagebox.showerror("오류", f"터미널 코드 '{terminal_code}'에 대한 매핑을 찾을 수 없습니다.")
                return

            # POL, TOL 설정
            self.selected_pol.set(port_info['pol'])
            self.selected_tol.set(port_info['tol'])

            # POL 버튼 색상 업데이트
            for port, btn in self.pol_buttons.items():
                if port == port_info['pol']:
                    btn.configure(bg='yellow')
                else:
                    btn.configure(bg='SystemButtonFace')

            # TOL 버튼 색상 업데이트
            for terminal, btn in self.tol_buttons.items():
                if terminal == port_info['tol']:
                    btn.configure(bg='yellow')
                else:
                    btn.configure(bg='SystemButtonFace')

            # 파일 정보 업데이트
            self.current_file = file_path
            self.input_label.config(text=f"입력 파일: {os.path.basename(file_path)}")
            
            # Summary 업데이트
            self.update_single_summary(df)
            
            # 멀티 탭의 Summary 초기화
            if hasattr(self, 'multi_summary_text'):
                self.multi_summary_text.delete(1.0, tk.END)
                self.multi_summary_text.insert(tk.END, "Multi CLL 탭에서 파일 병합 시 Summary가 표시됩니다.")
            
            # 선택된 서비스로 파일 변환
            self.convert_file(selected_service)
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다:\n{str(e)}")

    def convert_file(self, selected_service=None):
        """파일 변환"""
        try:
            if not selected_service:
                messagebox.showwarning("경고", "서비스를 선택해주세요!")
                return

            # CLL 파일 읽기
            cll_df = pd.read_excel(self.current_file, header=4)
            
            # 선택된 서비스의 매핑 가져오기
            service_mappings = self.stow_mapping.get(selected_service, [])
            
            # OBL 데이터프레임 생성
            obl_data = []

            # CLL 데이터 변환
            for idx, row in cll_df.iterrows():
                # OPT가 비어있으면 선택된 POL 값 사용
                por_value = row['OPT'] if pd.notna(row['OPT']) and row['OPT'] != '' else self.selected_pol.get()

                # POD와 FPOD 처리
                pod = str(row['POD']) if pd.notna(row['POD']) else ''
                fpod = str(row['FDP']) if pd.notna(row['FDP']) else ''
                
                # 초기값 설정
                mapped_port = pod
                mapped_stow = ''
                
                # POD가 stow_code와 일치하는지 확인
                for mapping in service_mappings:
                    if pod.upper() == mapping['stow_code'].upper():
                        mapped_port = mapping['port']
                        mapped_stow = mapping['stow_code']
                        break
                
                # OBL 데이터 생성
                obl_row = {
                    'No': idx + 1,
                    'CtrNbr': row['CNTR NO'],
                    'ShOwn': 'N',
                    'Opr': 'MSC',
                    'POR': por_value,
                    'POL': self.selected_pol.get(),
                    'TOL': self.selected_tol.get(),
                    'POD': mapped_port,
                    'TOD': '',
                    'Stow': mapped_stow,
                    'FPOD': fpod,
                    'SzTp': int(row['T&S']) if pd.notna(row['T&S']) else '',
                    'Wgt': int(row['WGT']) if pd.notna(row['WGT']) else '',
                    'ForE': row['F/E'],
                    'Rfopr': 'N',
                    'Door': 'C',
                    'CustH': 'N',
                    'Fumi': 'N',
                    'VGM': 'Y'
                }
                obl_data.append(obl_row)

            # OBL 데이터프레임 생성 및 저장
            obl_df = pd.DataFrame(obl_data)
            input_dir = os.path.dirname(self.current_file)
            base_name = os.path.splitext(os.path.basename(self.current_file))[0]
            output_file = os.path.join(input_dir, f"{base_name}_OBL.xlsx")
            obl_df.to_excel(output_file, index=False)

            self.output_file = output_file
            self.output_label.config(text=f"출력 파일: {output_file}")

            messagebox.showinfo("성공", "변환이 완료되었습니다.")
            
        except Exception as e:
            messagebox.showerror("Error", f"변환 중 오류 발생: {str(e)}")

    def drop_obl_file(self, event):
        """OBL 파일 드롭 처리"""
        file_path = event.data.strip('{}')
        self.current_file = file_path

        # 파일 정보 표시 업데이트
        self.input_label.config(text=f"입력 파일: {file_path}")
        self.obl_label.config(text=f"선택된 파일: {os.path.basename(file_path)}")

        # EMPTY 컨테이너 추가 실행
        self.add_empty_to_obl()

    def add_empty_to_obl(self):
        """기존 OBL에 EMPTY 컨테이너 추가"""
        # OBL 파일 읽기
        obl_df = pd.read_excel(self.current_file)

        # 기존 OBL의 컬럼 목록 가져오기
        existing_columns = obl_df.columns.tolist()

        # EMPTY 컨테이너 추가
        new_rows = []
        empty_container_num = 1  # 컨테이너 번호 시작값
        
        for pod_entry, sztp_entry, qty_entry in self.empty_entries:
            pod = pod_entry.get()
            sztp = sztp_entry.get()
            qty = qty_entry.get()

            if pod not in ["POD", ""] and sztp not in ["SzTp", ""] and qty not in ["수량", ""]:
                try:
                    qty = int(qty)
                    # SzTp를 정수로 변환
                    sztp = int(sztp)
                    
                    # SzTp에 따른 무게 설정
                    if str(sztp).startswith('2'):
                        weight = 2500
                    elif str(sztp).startswith('4'):
                        weight = 4500
                    else:
                        weight = 0

                    for i in range(qty):
                        # 기존 컬럼 구조를 따르는 빈 딕셔너리 생성
                        empty_row = {col: '' for col in existing_columns}

                        # 마지막 No 값 계산
                        last_no = len(obl_df) + len(new_rows) + 1

                        # EMPTY 컨테이너 번호 생성
                        ctr_nbr = f"MSCU{empty_container_num:07d}"
                        empty_container_num += 1

                        # 필요한 필드만 업데이트
                        empty_row.update({
                            'No': last_no,
                            'CtrNbr': ctr_nbr,  # 컨테이너 번호 설정
                            'ShOwn': 'N',
                            'Opr': 'MSC',
                            'POR': self.selected_pol.get(),
                            'POL': self.selected_pol.get(),
                            'TOL': self.selected_tol.get(),
                            'POD': pod,
                            'FPOD': pod,  # POD와 동일한 값으로 설정
                            'SzTp': sztp,
                            'Wgt': weight,  # SzTp에 따른 무게 설정
                            'ForE': 'E',
                            'Rfopr': 'N',
                            'Door': 'C',
                            'CustH': 'N',
                            'Fumi': 'N',
                            'VGM': 'Y',
                            'Stow': self.stow_mapping.get(pod, '')  # FPOD(POD)에 대한 Stow 코드
                        })
                        new_rows.append(empty_row)
                except ValueError:
                    continue  # 잘못된 입력은 조용히 건너뛰기

        # 새로운 EMPTY 컨테이너 추가
        if new_rows:
            new_df = pd.DataFrame(new_rows, columns=existing_columns)
            obl_df = pd.concat([obl_df, new_df], ignore_index=True)

            # 파일 저장
            input_dir = os.path.dirname(self.current_file)
            base_name = os.path.splitext(os.path.basename(self.current_file))[0]
            output_file = os.path.join(input_dir, f"{base_name}_EMPTY_ADDED.xlsx")
            obl_df.to_excel(output_file, index=False)

            self.output_file = output_file
            self.output_label.config(text=f"출력 파일: {output_file}")

            # Summary 업데이트
            self.update_summary(obl_df)

            messagebox.showinfo("성공", "EMPTY 컨테이너가 추가되었습니다.")

    def update_summary(self, df):
        """컨테이너 요약 정보 업데이트"""
        summary = "=== FULL 컨테이너 ===\n"
        full_containers = df[df['F/E'] == 'F']
        full_summary = full_containers['T&S'].value_counts()
        for sztp, count in full_summary.items():
            summary += f"{sztp}: {count}개\n"
        summary += f"FULL 컨테이너 총계: {len(full_containers)}개\n"

        summary += "\n=== EMPTY 컨테이너 ===\n"
        empty_containers = df[df['F/E'] == 'E']
        empty_summary = empty_containers['T&S'].value_counts()
        for sztp, count in empty_summary.items():
            summary += f"{sztp}: {count}개\n"

        # EMPTY 입력란에서 추가될 컨테이너 계산
        additional_empty = 0
        for pod_entry, sztp_entry, qty_entry in self.empty_entries:
            qty = qty_entry.get()
            if qty not in ["수량", ""]:
                try:
                    additional_empty += int(qty)
                except ValueError:
                    pass

        total_empty = len(empty_containers) + additional_empty
        summary += f"EMPTY 컨테이너 총계: {total_empty}개\n"

        # 전체 총계
        summary += f"\n=== 전체 컨테이너 ===\n"
        summary += f"총계: {len(full_containers) + total_empty}개"

        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(tk.END, summary)

    def select_pol(self, port):
        """POL 버튼 선택 처리"""
        self.selected_pol.set(port)
        # 모든 버튼 원래 색으로
        for btn in self.pol_buttons.values():
            btn.configure(bg='SystemButtonFace')
        # 선택된 버튼만 노란색으로
        self.pol_buttons[port].configure(bg='yellow')

    def select_tol(self, terminal):
        """TOL 버튼 선택 처리"""
        self.selected_tol.set(terminal)
        # 모든 버튼 원래 색으로
        for btn in self.tol_buttons.values():
            btn.configure(bg='SystemButtonFace')
        # 선택된 버튼만 노란색으로
        self.tol_buttons[terminal].configure(bg='yellow')

    def drop_itps_file(self, event):
        """ITPS 파일 드롭 처리"""
        file_path = event.data.strip('{}').strip('"')
        if not os.path.exists(file_path):
            messagebox.showerror("오류", "파일이 존재하지 않습니다.")
            return

        self.itps_file = file_path
        self.itps_input_label.config(text=f"ITPS 파일: {os.path.basename(file_path)}")
        self.itps_drop_label.config(text="ITPS 파일이 선택되었습니다")
        
        # 두 파일이 모두 선택되었다면 자동으로 처리 시작
        if self.itps_file and self.obl_file:
            self.process_itps_file()

    def drop_obl_for_itps(self, event):
        """ITPS 처리를 위한 OBL 파일 드롭 처리"""
        file_path = event.data.strip('{}').strip('"')
        if not os.path.exists(file_path):
            messagebox.showerror("오류", "파일이 존재하지 않습니다.")
            return

        self.obl_file = file_path
        self.itps_obl_label.config(text=f"OBL 파일: {os.path.basename(file_path)}")
        self.obl_drop_label.config(text="OBL 파일이 선택되었습니다")
        
        # 두 파일이 모두 선택되었다면 자동으로 처리 시작
        if self.itps_file and self.obl_file:
            self.process_itps_file()

    def process_itps_file(self):
        """ITPS 파일 처리 및 OBL에 추가"""
        try:
            # 선택된 서비스 확인
            selected_service = self.selected_service.get()
            if not selected_service:
                messagebox.showwarning("경고", "Service Name을 선택해주세요!")
                return

            # ITPS 파일 읽기 (헤더는 1행, 데이터는 3행부터)
            itps_df = pd.read_excel(self.itps_file, header=0, skiprows=[1])
            
            # OBL 파일 읽기
            obl_df = pd.read_excel(self.obl_file)
            
            # 기존 OBL의 마지막 No 값 가져오기
            last_no = obl_df['No'].max()
            
            # OBL의 POL과 TOL 값 가져오기
            obl_pol = obl_df['POL'].iloc[0] if not obl_df.empty else ''
            obl_tol = obl_df['TOL'].iloc[0] if not obl_df.empty else ''
            
            # 선택된 서비스의 매핑 가져오기
            service_mappings = self.stow_mapping.get(selected_service, [])
            
            # 기존 OBL 데이터에 대한 Stow Code 매핑 적용
            updated_obl_rows = []
            for _, row in obl_df.iterrows():
                obl_row = row.copy()
                pod = str(row['POD']) if pd.notna(row['POD']) else ''
                fpod = str(row['FPOD']) if pd.notna(row['FPOD']) else ''
                
                # POD가 stow_code와 일치하는지 확인
                mapped_port = pod
                mapped_stow = ''
                for mapping in service_mappings:
                    if pod.upper() == mapping['stow_code'].upper():
                        mapped_port = mapping['port']
                        mapped_stow = mapping['stow_code']
                        break
                
                obl_row['POD'] = mapped_port
                obl_row['Stow'] = mapped_stow
                obl_row['FPOD'] = fpod  # FPOD는 원래 값 유지
                updated_obl_rows.append(obl_row)
            
            # ITPS 데이터를 OBL 형식으로 변환
            new_rows = []
            for idx, row in itps_df.iterrows():
                try:
                    if pd.isna(row['Equipment Number']):
                        continue
                    
                    obl_row = {col: '' for col in obl_df.columns}
                    
                    # PORT CODE 변환 적용
                    por = self.convert_to_port_code(row['Origin Load Port']) if pd.notna(row['Origin Load Port']) else ''
                    pol = self.convert_to_port_code(obl_pol)  # OBL의 POL 사용
                    
                    # POD 값 가져오기
                    pod = str(row['Discharge Port']) if pd.notna(row['Discharge Port']) else ''
                    
                    # 초기값 설정
                    mapped_port = pod
                    mapped_stow = ''
                    
                    # POD가 stow_code와 일치하는지 확인
                    for mapping in service_mappings:
                        if pod.upper() == mapping['stow_code'].upper():
                            mapped_port = mapping['port']
                            mapped_stow = mapping['stow_code']
                            break
                    
                    # TpSZ 매핑 적용
                    tpsz = str(row['Type/Size']) if pd.notna(row['Type/Size']) else ''
                    mapped_tpsz = self.tpsz_mapping.get(tpsz, tpsz)
                    
                    # Rftemp 처리
                    rftemp = None
                    if pd.notna(row['Reefer Temp.']):
                        temp_str = str(row['Reefer Temp.']).split('/')[0].strip()
                        try:
                            rftemp = float(temp_str)
                        except ValueError:
                            rftemp = None
                    
                    # 나머지 필드 처리
                    obl_row.update({
                        'No': last_no + len(new_rows) + 1,
                        'CtrNbr': str(row['Equipment Number']) if pd.notna(row['Equipment Number']) else '',
                        'ShOwn': 'N',
                        'Opr': 'MSC',
                        'POR': por,
                        'POL': pol,
                        'TOL': obl_tol,
                        'POD': mapped_port,
                        'FPOD': pod,  # FPOD는 원래 값 유지
                        'Stow': mapped_stow,
                        'SzTp': mapped_tpsz,
                        'Wgt': int(row['Weight']) if pd.notna(row['Weight']) else '',
                        'ForE': str(row['Full/Empty']) if pd.notna(row['Full/Empty']) else 'N',
                        'Rfopr': 'N',
                        'Rftemp': f"{rftemp:.1f}" if rftemp is not None else '',
                        'Door': 'C',
                        'CustH': 'N',
                        'Fumi': 'N',
                        'VGM': 'Y',
                        'Class': str(int(row['IMO Class'])) if pd.notna(row['IMO Class']) and str(row['IMO Class']).replace('.', '').isdigit() else str(row['IMO Class']) if pd.notna(row['IMO Class']) else '',
                        'UNNO': str(row['UN Number'])[:6] if pd.notna(row['UN Number']) else ''
                    })
                    new_rows.append(obl_row)
                except Exception as e:
                    print(f"행 {idx} 데이터 확인 중 오류: {str(e)}")
                    continue
            
            # 기존 OBL 데이터와 새로운 ITPS 데이터 결합
            updated_obl_df = pd.DataFrame(updated_obl_rows)
            new_df = pd.DataFrame(new_rows)
            combined_df = pd.concat([updated_obl_df, new_df], ignore_index=True)
            
            # 모든 port 코드 변환 적용
            combined_df['POR'] = combined_df['POR'].apply(self.convert_to_port_code)
            combined_df['POL'] = combined_df['POL'].apply(self.convert_to_port_code)
            combined_df['POD'] = combined_df['POD'].apply(self.convert_to_port_code)
            combined_df['FPOD'] = combined_df['FPOD'].apply(self.convert_to_port_code)
            
            # 파일 저장
            save_dir = os.path.dirname(self.obl_file)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(save_dir, f"OBL_with_ITPS_{timestamp}.xlsx")
            combined_df.to_excel(output_file, index=False)
            
            # 결과 표시
            self.itps_output_label.config(text=f"출력 파일: {os.path.basename(output_file)}")
            
            # Summary 업데이트
            self.update_itps_summary(combined_df)
            
            messagebox.showinfo("성공", "ITPS 데이터가 성공적으로 추가되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"ITPS 처리 중 오류 발생: {str(e)}")

    def update_itps_summary(self, df):
        """ITPS 처리 결과 Summary 업데이트"""
        try:
            self.itps_summary_text.delete(1.0, tk.END)
            
            summary_text = "=== ITPS 추가 결과 Summary ===\n"
            summary_text += "================================\n\n"
            
            # 전체 컨테이너 수
            total_containers = len(df)
            summary_text += f"전체 컨테이너 수: {total_containers}\n"
            summary_text += "--------------------------------\n\n"
            
            # F/E 별 통계
            fe_counts = df['ForE'].value_counts()
            summary_text += "=== Full/Empty 현황 ===\n"
            for fe, count in fe_counts.items():
                summary_text += f"{fe}: {count}개\n"
            summary_text += "--------------------------------\n\n"
            
            # Size Type 별 통계
            sztp_counts = df['SzTp'].value_counts()
            summary_text += "=== Size Type 현황 ===\n"
            for sztp, count in sztp_counts.items():
                if pd.notna(sztp):
                    summary_text += f"{sztp}: {count}개\n"
            summary_text += "--------------------------------\n\n"
            
            # POD 별 통계
            pod_counts = df['POD'].value_counts()
            summary_text += "=== POD 현황 ===\n"
            for pod, count in pod_counts.items():
                if pd.notna(pod):
                    summary_text += f"{pod}: {count}개\n"
            summary_text += "--------------------------------"
            
            self.itps_summary_text.insert(tk.END, summary_text)
            
        except Exception as e:
            self.itps_summary_text.delete(1.0, tk.END)
            self.itps_summary_text.insert(tk.END, f"Summary 생성 중 오류 발생: {str(e)}")

    def convert_to_port_code(self, port_name):
        """항구 이름을 5자리 PORT CODE로 변환"""
        if not port_name or pd.isna(port_name):
            return ''
            
        port_name = str(port_name).strip().upper()
        
        # 이미 5자리 코드인 경우 그대로 반환
        if len(port_name) == 5 and port_name.isalnum():
            return port_name
            
        # port_codes의 value(port name)와 매칭 시도
        for code, full_name in self.port_codes.items():
            if full_name == port_name:  # 정확한 매칭
                return code
            elif full_name in port_name or port_name in full_name:  # 부분 매칭
                return code
                
        # 매칭되는 코드가 없으면 원래 값 반환
        return port_name

    def drop_stowage_file(self, event):
        """Stowage Code 엑셀 파일 드롭 처리"""
        try:
            file_path = event.data.strip('{}').strip('"')
            if not os.path.exists(file_path):
                messagebox.showerror("오류", "파일이 존재하지 않습니다.")
                return

            # 엑셀 파일 읽기 (헤더는 2번째 행, 데이터는 3번째 행부터)
            df = pd.read_excel(file_path, header=1)
            
            # 매핑 딕셔너리 생성
            service_mappings = {}
            for _, row in df.iterrows():
                service_name = str(row['Service Name']).strip()
                stow_code = str(row['Stow Code OBL7']).strip()
                
                # Port 열에서 [ ] 안의 값 추출
                port_str = str(row['Port']).strip()
                port = ''
                if '[' in port_str and ']' in port_str:
                    start = port_str.find('[') + 1
                    end = port_str.find(']')
                    port = port_str[start:end].strip()
                
                if port and stow_code:  # port와 stow_code가 있는 경우만 매핑에 추가
                    if service_name not in service_mappings:
                        service_mappings[service_name] = []
                    service_mappings[service_name].append({
                        'port': port,
                        'stow_code': stow_code
                    })
            
            # 설정 저장
            self.stow_mapping = service_mappings
            
            # 엑셀 파일 경로 저장
            excel_dir = os.path.dirname(file_path)
            excel_name = os.path.splitext(os.path.basename(file_path))[0]
            self.stowage_config_file = os.path.join(excel_dir, f"{excel_name}_mapping.json")
            
            self.save_stowage_settings()
            
            # 미리보기 업데이트
            self.update_stowage_preview()
            
            messagebox.showinfo("성공", "Stowage Code 매핑이 성공적으로 업데이트되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다: {str(e)}")

    def save_stowage_settings(self):
        """Stowage Code 설정 저장"""
        try:
            # JSON 파일로 저장
            with open(self.stowage_config_file, 'w', encoding='utf-8') as f:
                json.dump(self.stow_mapping, f, ensure_ascii=False, indent=2)
                
            messagebox.showinfo("성공", f"설정이 성공적으로 저장되었습니다.\n저장 위치: {self.stowage_config_file}")
            
        except Exception as e:
            messagebox.showerror("오류", f"설정 저장 중 오류가 발생했습니다: {str(e)}")

    def update_stowage_preview(self):
        """Stowage Code 매핑 미리보기 업데이트"""
        try:
            self.preview_text.delete(1.0, tk.END)
            
            # 서비스 목록 업데이트
            service_names = list(self.stow_mapping.keys())
            self.service_combo['values'] = service_names
            
            preview_text = "=== 현재 매핑 ===\n"
            selected_service = self.selected_service.get()
            
            if selected_service:
                preview_text += f"Service Name: {selected_service}\n"
                preview_text += "------------------------\n"
                
                # 선택된 서비스에 대한 매핑만 표시
                if selected_service in self.stow_mapping:
                    for mapping in self.stow_mapping[selected_service]:
                        preview_text += f"POD: {mapping['port']}\n"
                        preview_text += f"Stow Code: {mapping['stow_code']}\n"
                        preview_text += "------------------------\n"
            else:
                # 서비스가 선택되지 않은 경우 모든 매핑 표시
                for service_name, mappings in self.stow_mapping.items():
                    preview_text += f"Service Name: {service_name}\n"
                    for mapping in mappings:
                        preview_text += f"POD: {mapping['port']}\n"
                        preview_text += f"Stow Code: {mapping['stow_code']}\n"
                    preview_text += "------------------------\n"
                
            self.preview_text.insert(tk.END, preview_text)
            
        except Exception as e:
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, f"미리보기 업데이트 중 오류 발생: {str(e)}")

    def drop_tpsz_file(self, event):
        """TpSZ 엑셀 파일 드롭 처리"""
        try:
            file_path = event.data.strip('{}').strip('"')
            if not os.path.exists(file_path):
                messagebox.showerror("오류", "파일이 존재하지 않습니다.")
                return

            # 엑셀 파일 읽기
            df = pd.read_excel(file_path)
            
            # 컬럼 매핑 가져오기
            before_col = self.before_entry.get().strip()
            after_col = self.after_entry.get().strip()
            
            if not before_col or not after_col:
                messagebox.showerror("오류", "컬럼 매핑을 먼저 설정해주세요.")
                return
                
            if before_col not in df.columns or after_col not in df.columns:
                messagebox.showerror("오류", "설정한 컬럼명이 엑셀 파일에 존재하지 않습니다.")
                return

            # 매핑 딕셔너리 생성
            mapping = dict(zip(df[before_col], df[after_col]))
            
            # 설정 저장
            self.tpsz_mapping = mapping
            
            # JSON 파일 경로를 엑셀 파일과 동일한 디렉토리로 설정
            excel_dir = os.path.dirname(file_path)
            excel_name = os.path.splitext(os.path.basename(file_path))[0]
            self.tpsz_config_file = os.path.join(excel_dir, f"{excel_name}_mapping.json")
            
            self.save_tpsz_settings()
            
            # 미리보기 업데이트
            self.update_tpsz_preview()
            
            messagebox.showinfo("성공", "TpSZ 매핑이 성공적으로 업데이트되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다: {str(e)}")

    def save_tpsz_settings(self):
        """TpSZ 설정 저장"""
        try:
            # 현재 설정 가져오기
            settings = {
                'column_mapping': {
                    'before': self.before_entry.get().strip(),
                    'after': self.after_entry.get().strip()
                },
                'mapping': self.tpsz_mapping
            }
            
            # JSON 파일로 저장
            with open(self.tpsz_config_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
                
            messagebox.showinfo("성공", f"설정이 성공적으로 저장되었습니다.\n저장 위치: {self.tpsz_config_file}")
            
        except Exception as e:
            messagebox.showerror("오류", f"설정 저장 중 오류가 발생했습니다: {str(e)}")

    def update_tpsz_preview(self):
        """TpSZ 매핑 미리보기 업데이트"""
        try:
            self.tpsz_preview_text.delete(1.0, tk.END)
            
            preview_text = "=== 컬럼 매핑 설정 ===\n"
            preview_text += f"Before 컬럼: {self.tpsz_column_mapping.get('before', '')}\n"
            preview_text += f"After 컬럼: {self.tpsz_column_mapping.get('after', '')}\n\n"
            
            preview_text += "=== 현재 매핑 ===\n"
            for before, after in self.tpsz_mapping.items():
                preview_text += f"{before}: {after}\n"
                
            self.tpsz_preview_text.insert(tk.END, preview_text)
            
        except Exception as e:
            self.tpsz_preview_text.delete(1.0, tk.END)
            self.tpsz_preview_text.insert(tk.END, f"미리보기 업데이트 중 오류 발생: {str(e)}")

    def setup_tpsz_tab(self):
        """TpSZ 관리 탭 설정"""
        # 메인 프레임
        main_frame = ttk.Frame(self.tpsz_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 드래그 & 드롭 영역
        drop_frame = ttk.LabelFrame(main_frame, text="TpSZ 엑셀 파일")
        drop_frame.pack(fill="x", pady=(0, 10))

        self.tpsz_drop_label = ttk.Label(drop_frame, text="TpSZ 엑셀 파일을 여기에 드롭하세요")
        self.tpsz_drop_label.pack(pady=20)

        # 드래그 앤 드롭 바인딩
        self.tpsz_drop_label.drop_target_register(DND_FILES)
        self.tpsz_drop_label.dnd_bind('<<Drop>>', self.drop_tpsz_file)

        # 컬럼 매핑 설정 영역
        mapping_frame = ttk.LabelFrame(main_frame, text="컬럼 매핑 설정")
        mapping_frame.pack(fill="x", pady=(0, 10))

        # Before 컬럼 매핑
        before_frame = ttk.Frame(mapping_frame)
        before_frame.pack(fill="x", pady=5)
        ttk.Label(before_frame, text="Before 컬럼명:").pack(side="left", padx=5)
        self.before_entry = ttk.Entry(before_frame)
        self.before_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.before_entry.insert(0, self.tpsz_column_mapping.get('before', ''))

        # After 컬럼 매핑
        after_frame = ttk.Frame(mapping_frame)
        after_frame.pack(fill="x", pady=5)
        ttk.Label(after_frame, text="After 컬럼명:").pack(side="left", padx=5)
        self.after_entry = ttk.Entry(after_frame)
        self.after_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.after_entry.insert(0, self.tpsz_column_mapping.get('after', ''))

        # 저장 버튼
        save_button = ttk.Button(mapping_frame, text="설정 저장", command=self.save_tpsz_settings)
        save_button.pack(pady=10)

        # 현재 매핑 미리보기
        preview_frame = ttk.LabelFrame(main_frame, text="현재 매핑 미리보기")
        preview_frame.pack(fill="both", expand=True)
        
        self.tpsz_preview_text = tk.Text(preview_frame, height=10)
        self.tpsz_preview_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 현재 매핑 표시
        self.update_tpsz_preview()

    def terminal_to_port_mapping(self, terminal_code):
        # 터미널 코드에 따른 POL, TOL 매핑 딕셔너리
        terminal_mapping = {
            'PNITC': {'pol': 'KRPUS', 'tol': 'KRPUSAB'},
            'PNCOC': {'pol': 'KRPUS', 'tol': 'KRPUSPN'},
            'BCTHD': {'pol': 'KRPUS', 'tol': 'KRPUSBC'},
            'HJNPC': {'pol': 'KRPUS', 'tol': 'KRPUSAP'},
            'ICTPC': {'pol': 'KRINCAH', 'tol': 'KRINCAH'},
            'KEGWC': {'pol': 'KRKAN', 'tol': 'KRKANKT'}
        }
        
        return terminal_mapping.get(terminal_code, {'pol': '', 'tol': ''})

    def process_cll_file(self):
        # ... existing code ...
        
        # 엑셀 파일에서 (12,4) 위치의 터미널 코드 읽기
        terminal_code = worksheet.cell(12, 4).value
        
        # 터미널 코드를 기반으로 POL, TOL 값 설정
        port_info = self.terminal_to_port_mapping(terminal_code)
        self.pol_value = port_info['pol']
        self.tol_value = port_info['tol']
        
        # ... existing code ...

    def reset_all(self):
        """프로그램 상태 초기화"""
        # POL/TOL 버튼 초기화 (단일 탭)
        for btn in self.pol_buttons.values():
            btn.configure(bg='SystemButtonFace')
        for btn in self.tol_buttons.values():
            btn.configure(bg='SystemButtonFace')
        
        # POL/TOL 버튼 초기화 (멀티 탭)
        for btn in self.multi_pol_buttons.values():
            btn.configure(bg='SystemButtonFace')
        for btn in self.multi_tol_buttons.values():
            btn.configure(bg='SystemButtonFace')
        
        # 선택값 초기화
        self.selected_pol.set('')
        self.selected_tol.set('')
        
        # 파일 경로 레이블 초기화
        self.input_label.config(text="입력 파일: 없음")
        self.output_label.config(text="출력 파일: 없음")
        self.master_path_label.config(text="파일 경로: 없음")
        self.slave_path_label.config(text="파일 경로: 없음")
        self.result_label.config(text="출력 파일: 없음")
        
        # Summary 텍스트 초기화
        self.single_summary_text.delete(1.0, tk.END)
        self.single_summary_text.insert(tk.END, "단일 CLL 탭에서 파일 변환 시 Summary가 표시됩니다.")
        self.multi_summary_text.delete(1.0, tk.END)
        self.multi_summary_text.insert(tk.END, "Multi CLL 탭에서 파일 변환 시 Summary가 표시됩니다.")
        
        # 파일 관련 변수 초기화
        self.current_file = None
        self.output_file = None
        if hasattr(self, 'master_file'):
            delattr(self, 'master_file')
        if hasattr(self, 'slave_file'):
            delattr(self, 'slave_file')

        # Entry 위젯 초기화
        for pod_entry, sztp_entry, qty_entry in self.empty_entries:
            # Entry 위젯 상태 초기화
            pod_entry.delete(0, tk.END)
            sztp_entry.delete(0, tk.END)
            qty_entry.delete(0, tk.END)
            
            # 플레이스홀더 텍스트 설정
            pod_entry.insert(0, "POD")
            sztp_entry.insert(0, "SzTp")
            qty_entry.insert(0, "수량")
            
            # Entry 위젯 상태 설정
            pod_entry.config(state='normal')
            sztp_entry.config(state='normal')
            qty_entry.config(state='normal')

    def run(self):
        self.root.mainloop()

    def drop_master_cll(self, event):
        """Master CLL 파일 드롭 처리"""
        try:
            file_path = event.data.strip('{}').strip('"')
            if not os.path.exists(file_path):
                messagebox.showerror("오류", "파일이 존재하지 않습니다.")
                return

            # 엑셀 파일에서 (4,12) 위치의 터미널 코드 읽기
            df_check = pd.read_excel(file_path, header=None)
            terminal_code = str(df_check.iloc[3, 11]).strip()

            if not terminal_code:
                messagebox.showerror("오류", "(4,12) 위치에서 터미널 코드를 찾을 수 없습니다.")
                return

            # 터미널 코드를 기반으로 POL, TOL 값 자동 설정
            port_info = self.terminal_to_port_mapping(terminal_code)
            
            if not port_info['pol'] or not port_info['tol']:
                messagebox.showerror("오류", f"터미널 코드 '{terminal_code}'에 대한 매핑을 찾을 수 없습니다.")
                return

            # POL, TOL 설정
            self.selected_pol.set(port_info['pol'])
            self.selected_tol.set(port_info['tol'])

            # POL 버튼 색상 업데이트 (멀티 탭)
            for port, btn in self.multi_pol_buttons.items():
                if port == port_info['pol']:
                    btn.configure(bg='yellow')
                else:
                    btn.configure(bg='SystemButtonFace')

            # TOL 버튼 색상 업데이트 (멀티 탭)
            for terminal, btn in self.multi_tol_buttons.items():
                if terminal == port_info['tol']:
                    btn.configure(bg='yellow')
                else:
                    btn.configure(bg='SystemButtonFace')

            self.master_file = file_path
            self.master_path_label.config(text=f"파일 경로: {os.path.basename(file_path)}")
            self.master_label.config(text="Master 파일이 선택되었습니다")
            
            # POD 목록 추출
            df = pd.read_excel(file_path, header=4)
            pod_list = df['POD'].unique().tolist()
            
            # Master 파일의 POD 목록과 합치기
            master_df = pd.read_excel(self.master_file, header=4)
            master_pods = master_df['POD'].unique().tolist()
            all_pods = list(set(pod_list + master_pods))
            
            # 매칭되는 서비스 찾기 (아직 서비스가 선택되지 않은 경우)
            if not hasattr(self, 'selected_service'):
                matching_services = self.find_matching_services(all_pods)
                if matching_services:
                    selected_service = self.show_service_selection_dialog(matching_services)
                    if selected_service:
                        self.selected_service = selected_service
                else:
                    messagebox.showwarning("경고", "POD와 일치하는 서비스를 찾을 수 없습니다.")
                    return

            # Slave 프레임 활성화
            self.slave_frame.pack(pady=10, padx=10, fill="x")

        except Exception as e:
            error_msg = str(e)
            print(f"Error in drop_master_cll: {error_msg}")  # 디버깅용
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다:\n{error_msg}")

    def drop_slave_cll(self, event):
        """Slave CLL 파일 드롭 처리"""
        try:
            if not hasattr(self, 'master_file'):
                messagebox.showwarning("경고", "Master 파일을 먼저 선택해주세요!")
                return

            file_path = event.data.strip('{}').strip('"')
            if not os.path.exists(file_path):
                messagebox.showerror("오류", "파일이 존재하지 않습니다.")
                return

            # 엑셀 파일에서 (4,12) 위치의 터미널 코드 읽기
            df_check = pd.read_excel(file_path, header=None)
            terminal_code = str(df_check.iloc[3, 11]).strip()

            if not terminal_code:
                messagebox.showerror("오류", "(4,12) 위치에서 터미널 코드를 찾을 수 없습니다.")
                return

            # 터미널 코드 확인
            port_info = self.terminal_to_port_mapping(terminal_code)
            if port_info['pol'] != self.selected_pol.get() or port_info['tol'] != self.selected_tol.get():
                messagebox.showerror("오류", "Master 파일과 POL/TOL이 일치하지 않습니다.")
                return

            self.slave_file = file_path
            self.slave_path_label.config(text=f"파일 경로: {os.path.basename(file_path)}")
            self.slave_label.config(text="Slave 파일이 선택되었습니다")
            
            # POD 목록 추출
            df = pd.read_excel(file_path, header=4)
            pod_list = df['POD'].unique().tolist()
            
            # Master 파일의 POD 목록과 합치기
            master_df = pd.read_excel(self.master_file, header=4)
            master_pods = master_df['POD'].unique().tolist()
            all_pods = list(set(pod_list + master_pods))
            
            # 매칭되는 서비스 찾기 (아직 서비스가 선택되지 않은 경우)
            if not hasattr(self, 'selected_service'):
                matching_services = self.find_matching_services(all_pods)
                if matching_services:
                    selected_service = self.show_service_selection_dialog(matching_services)
                    if selected_service:
                        self.selected_service = selected_service
                else:
                    messagebox.showwarning("경고", "POD와 일치하는 서비스를 찾을 수 없습니다.")
                    return

            # 파일 병합 처리 시작
            self.combine_cll_files()

        except Exception as e:
            error_msg = str(e)
            print(f"Error in drop_slave_cll: {error_msg}")  # 디버깅용
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다:\n{error_msg}")

    def combine_cll_files(self):
        """Master와 Slave CLL 파일 병합"""
        try:
            # Master와 Slave 파일 읽기
            master_df = pd.read_excel(self.master_file, header=4)
            slave_df = pd.read_excel(self.slave_file, header=4)

            # 선택된 서비스의 매핑 가져오기
            service_mappings = self.stow_mapping.get(self.selected_service, [])

            # OBL 데이터프레임 생성
            obl_data = []

            # Master와 Slave 데이터 변환 및 병합
            for df in [master_df, slave_df]:
                for idx, row in df.iterrows():
                    # OPT가 비어있으면 선택된 POL 값 사용
                    por_value = row['OPT'] if pd.notna(row['OPT']) and row['OPT'] != '' else self.selected_pol.get()

                    # POD와 FPOD 처리
                    pod = str(row['POD']) if pd.notna(row['POD']) else ''
                    fpod = str(row['FDP']) if pd.notna(row['FDP']) else ''
                    
                    # 초기값 설정
                    mapped_port = pod
                    mapped_stow = ''
                    
                    # POD가 stow_code와 일치하는지 확인
                    for mapping in service_mappings:
                        if pod.upper() == mapping['stow_code'].upper():
                            mapped_port = mapping['port']
                            mapped_stow = mapping['stow_code']
                            break

                    # OBL 데이터 생성
                    obl_row = {
                        'No': len(obl_data) + 1,  # 연속된 번호 부여
                        'CtrNbr': row['CNTR NO'],
                        'ShOwn': 'N',
                        'Opr': 'MSC',
                        'POR': por_value,
                        'POL': self.selected_pol.get(),
                        'TOL': self.selected_tol.get(),
                        'POD': mapped_port,
                        'TOD': '',
                        'Stow': mapped_stow,
                        'FPOD': fpod,
                        'SzTp': int(row['T&S']) if pd.notna(row['T&S']) else '',
                        'Wgt': int(row['WGT']) if pd.notna(row['WGT']) else '',
                        'ForE': row['F/E'],
                        'Rfopr': 'N',
                        'Door': 'C',
                        'CustH': 'N',
                        'Fumi': 'N',
                        'VGM': 'Y'
                    }
                    obl_data.append(obl_row)

            # OBL 데이터프레임 생성 및 저장
            combined_df = pd.DataFrame(obl_data)
            output_dir = os.path.dirname(self.master_file)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_dir, f"Combined_OBL_{timestamp}.xlsx")
            combined_df.to_excel(output_file, index=False)

            # 결과 표시
            self.result_label.config(text=f"출력 파일: {os.path.basename(output_file)}")

            # Summary 업데이트
            self.update_multi_summary(combined_df)

            messagebox.showinfo("성공", "파일이 성공적으로 병합되었습니다.")

        except Exception as e:
            messagebox.showerror("오류", f"파일 병합 중 오류가 발생했습니다: {str(e)}")

    def update_multi_summary(self, df):
        """Multi CLL 탭의 Summary 업데이트"""
        try:
            summary_text = "=== Container Summary ===\n\n"
            
            # 전체 컨테이너 수
            total_containers = len(df)
            summary_text += f"전체 컨테이너 수: {total_containers}\n\n"
            
            # Size Type별 통계
            sztp_counts = df['SzTp'].value_counts()
            summary_text += "=== Size Type 현황 ===\n"
            for sztp, count in sztp_counts.items():
                if pd.notna(sztp):
                    summary_text += f"{sztp}: {count}개\n"
            
            # POD별 통계
            summary_text += "\n=== POD 현황 ===\n"
            pod_counts = df['POD'].value_counts()
            for pod, count in pod_counts.items():
                if pd.notna(pod):
                    summary_text += f"{pod}: {count}개\n"

            self.multi_summary_text.delete(1.0, tk.END)
            self.multi_summary_text.insert(tk.END, summary_text)

        except Exception as e:
            self.multi_summary_text.delete(1.0, tk.END)
            self.multi_summary_text.insert(tk.END, f"Summary 생성 중 오류 발생: {str(e)}")

    def update_single_summary(self, df):
        """단일 CLL 탭의 Summary 업데이트"""
        try:
            summary_text = "=== Container Summary ===\n\n"
            
            # 전체 컨테이너 수
            total_containers = len(df)
            summary_text += f"전체 컨테이너 수: {total_containers}\n\n"
            
            # Size Type별 통계
            sztp_counts = df['T&S'].value_counts()
            summary_text += "=== Size Type 현황 ===\n"
            for sztp, count in sztp_counts.items():
                if pd.notna(sztp):
                    summary_text += f"{sztp}: {count}개\n"
            
            # F/E 별 통계
            summary_text += "\n=== Full/Empty 현황 ===\n"
            fe_counts = df['F/E'].value_counts()
            for fe, count in fe_counts.items():
                if pd.notna(fe):
                    summary_text += f"{fe}: {count}개\n"
            
            # POD별 통계
            summary_text += "\n=== POD 현황 ===\n"
            pod_counts = df['POD'].value_counts()
            for pod, count in pod_counts.items():
                if pd.notna(pod):
                    summary_text += f"{pod}: {count}개\n"

            self.single_summary_text.delete(1.0, tk.END)
            self.single_summary_text.insert(tk.END, summary_text)

        except Exception as e:
            self.single_summary_text.delete(1.0, tk.END)
            self.single_summary_text.insert(tk.END, f"Summary 생성 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    app = ContainerConverter()
    app.run()


#test4