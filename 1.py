import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import random

class TitleInserterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("제목 자동 삽입 프로그램")
        self.root.geometry("800x800")
        
        # 데이터 저장용 변수들
        self.manuscript_files = []
        self.title_list = []
        self.title_file_path = ""
        self.tag_files = []
        self.tag_settings = []
        self.filter_words = []
        self.filter_file_path = ""
        self.save_path = ""
        self.apply_filter_var = tk.BooleanVar(value=False)
        self.random_tag_position_var = tk.BooleanVar(value=False)
        self.tag_count_var = tk.StringVar(value="5")
        self.tag_spacing_var = tk.StringVar(value="20")
        self.tag_file_count_var = tk.StringVar(value="3")
        
        # 제목 삽입 모드 설정
        self.title_mode_var = tk.StringVar(value="single")
        self.title_repeat_var = tk.BooleanVar()
        self.random_repeat_var = tk.BooleanVar()
        self.repeat_count_var = tk.StringVar(value="5")
        self.append_word_var = tk.BooleanVar()
        self.append_word_text = tk.StringVar()
        
        # 치환 기능 변수들 (치환 횟수 추가)
        self.replace_enabled_var = tk.BooleanVar(value=False)
        self.replace_target_var = tk.StringVar(value="장롱면허운전연수")
        self.replace_count_var = tk.StringVar(value="1")  # 치환 횟수 추가
        
        # 파일명에 제목 앞단어 추가 기능
        self.add_title_word_to_filename_var = tk.BooleanVar(value=False)
        
        # 태그 파일 로딩 모드 변수 추가
        self.tag_load_mode_var = tk.StringVar(value="individual")  # individual or batch
        
        self.create_widgets()
        self.update_tag_files()
    
    def create_widgets(self):
        # Canvas와 Scrollbar 추가
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.main_frame = ttk.Frame(self.canvas)
        
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        self.main_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        main_frame = self.main_frame
        main_frame.configure(padding="10")
        
        # 원고 파일 섹션
        manuscript_frame = ttk.LabelFrame(main_frame, text="원고 파일 관리", padding="10")
        manuscript_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(manuscript_frame, text="원고 파일 추가", command=self.add_manuscript_files).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(manuscript_frame, text="선택 파일 제거", command=self.remove_selected_file).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(manuscript_frame, text="전체 파일 삭제", command=self.clear_all_files).grid(row=0, column=2)
        
        self.file_listbox = tk.Listbox(manuscript_frame, height=6, width=80)
        self.file_listbox.grid(row=1, column=0, columnspan=3, pady=(10, 0), sticky=(tk.W, tk.E))
        
        # 제목 파일 섹션
        title_frame = ttk.LabelFrame(main_frame, text="제목 파일 관리", padding="10")
        title_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(title_frame, text="제목 파일 불러오기", command=self.load_title_file).grid(row=0, column=0, padx=(0, 10))
        self.title_file_label = ttk.Label(title_frame, text="제목 파일이 선택되지 않았습니다.")
        self.title_file_label.grid(row=0, column=1, sticky=tk.W)
        
        self.title_preview = scrolledtext.ScrolledText(title_frame, height=4, width=80, state=tk.DISABLED)
        self.title_preview.grid(row=1, column=0, columnspan=2, pady=(10, 0), sticky=(tk.W, tk.E))
        
        # 치환 기능 섹션 (치환 횟수 추가)
        replace_frame = ttk.LabelFrame(main_frame, text="치환 기능", padding="10")
        replace_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.replace_check = ttk.Checkbutton(replace_frame, text="치환 기능 사용", variable=self.replace_enabled_var)
        self.replace_check.grid(row=0, column=0, padx=(0, 20), sticky=tk.W)
        
        ttk.Label(replace_frame, text="치환 대상 문구:").grid(row=0, column=1, padx=(0, 10), sticky=tk.W)
        self.replace_target_entry = ttk.Entry(replace_frame, textvariable=self.replace_target_var, width=30)
        self.replace_target_entry.grid(row=0, column=2, sticky=tk.W, padx=(0, 10))
        
        ttk.Label(replace_frame, text="치환 횟수:").grid(row=0, column=3, padx=(0, 10), sticky=tk.W)
        self.replace_count_entry = ttk.Entry(replace_frame, textvariable=self.replace_count_var, width=10)
        self.replace_count_entry.grid(row=0, column=4, sticky=tk.W)
        
        ttk.Label(replace_frame, text="※ 제목 파일의 첫 번째 단어로 지정된 횟수만큼 치환됩니다", foreground="gray").grid(row=1, column=0, columnspan=5, pady=(5, 0), sticky=tk.W)
        
        # 태그 파일 섹션
        self.tag_frame = ttk.LabelFrame(main_frame, text="태그 파일 관리", padding="10")
        self.tag_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 태그 파일 로딩 모드 선택
        tag_load_mode_frame = ttk.Frame(self.tag_frame)
        tag_load_mode_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(tag_load_mode_frame, text="태그 파일 로딩 방식:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        ttk.Radiobutton(tag_load_mode_frame, text="개별 불러오기", variable=self.tag_load_mode_var, value="individual", command=self.on_tag_load_mode_change).grid(row=0, column=1, padx=(0, 10))
        ttk.Radiobutton(tag_load_mode_frame, text="한번에 불러오기", variable=self.tag_load_mode_var, value="batch", command=self.on_tag_load_mode_change).grid(row=0, column=2, padx=(0, 10))
        ttk.Button(tag_load_mode_frame, text="일괄 태그 파일 선택", command=self.load_batch_tag_files).grid(row=0, column=3, padx=(10, 0))
        
        # 태그 파일 개수 입력
        tag_count_frame = ttk.Frame(self.tag_frame)
        tag_count_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(tag_count_frame, text="태그 파일 개수 (1~50):").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        self.tag_file_count_entry = ttk.Entry(tag_count_frame, textvariable=self.tag_file_count_var, width=10)
        self.tag_file_count_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        ttk.Button(tag_count_frame, text="적용", command=self.update_tag_files).grid(row=0, column=2, padx=(0, 10))
        
        # 태그 파일 입력 위젯을 동적으로 생성하기 위한 프레임
        self.tag_files_frame = ttk.Frame(self.tag_frame)
        self.tag_files_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        # 필터링 파일 섹션
        filter_frame = ttk.LabelFrame(main_frame, text="필터링 파일 관리", padding="10")
        filter_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(filter_frame, text="필터링 파일 불러오기", command=self.load_filter_file).grid(row=0, column=0, padx=(0, 10))
        self.filter_file_label = ttk.Label(filter_frame, text="필터링 파일이 선택되지 않았습니다.")
        self.filter_file_label.grid(row=0, column=1, sticky=tk.W)
        
        self.apply_filter_check = ttk.Checkbutton(filter_frame, text="필터링 적용", variable=self.apply_filter_var)
        self.apply_filter_check.grid(row=1, column=0, columnspan=2, pady=(10, 0), sticky=(tk.W))
        
        # 저장 경로 설정 섹션
        save_path_frame = ttk.LabelFrame(main_frame, text="저장 경로 설정", padding="10")
        save_path_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        path_type_frame = ttk.Frame(save_path_frame)
        path_type_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.path_type_var = tk.StringVar(value="original")
        ttk.Radiobutton(path_type_frame, text="원본 폴더에 새파일 생성", variable=self.path_type_var, value="original").grid(row=0, column=0, padx=(0, 10))
        ttk.Radiobutton(path_type_frame, text="절대경로로 저장", variable=self.path_type_var, value="absolute").grid(row=0, column=1, padx=(0, 10))
        ttk.Radiobutton(path_type_frame, text="상대경로로 저장", variable=self.path_type_var, value="relative").grid(row=0, column=2)
        
        path_select_frame = ttk.Frame(save_path_frame)
        path_select_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(path_select_frame, text="저장 폴더 선택", command=self.select_save_path).grid(row=0, column=0, padx=(0, 10))
        self.save_path_label = ttk.Label(path_select_frame, text="저장 경로가 설정되지 않았습니다.")
        self.save_path_label.grid(row=0, column=1, sticky=tk.W)
        
        relative_path_frame = ttk.Frame(save_path_frame)
        relative_path_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        ttk.Label(relative_path_frame, text="상대경로 입력:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        self.relative_path_var = tk.StringVar(value="output")
        self.relative_path_entry = ttk.Entry(relative_path_frame, textvariable=self.relative_path_var, width=50)
        self.relative_path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 파일명에 제목 앞단어 추가 옵션
        filename_option_frame = ttk.Frame(save_path_frame)
        filename_option_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.add_title_word_check = ttk.Checkbutton(filename_option_frame, text="저장 시 제목 앞단어 추가", variable=self.add_title_word_to_filename_var)
        self.add_title_word_check.grid(row=0, column=0, sticky=tk.W)
        ttk.Label(filename_option_frame, text="※ 제목 파일의 첫 번째 단어를 파일명에 추가합니다", foreground="gray").grid(row=0, column=1, padx=(10, 0), sticky=tk.W)
        
        # 설정 섹션
        self.settings_frame = ttk.LabelFrame(main_frame, text="설정", padding="10")
        self.settings_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 제목 삽입 모드 선택
        title_mode_frame = ttk.Frame(self.settings_frame)
        title_mode_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(title_mode_frame, text="단일 제목 삽입", variable=self.title_mode_var, value="single").grid(row=0, column=0, padx=(0, 10))
        ttk.Radiobutton(title_mode_frame, text="여러 제목 삽입", variable=self.title_mode_var, value="multiple").grid(row=0, column=1, padx=(0, 10))
        
        # 제목 반복 설정
        repeat_frame = ttk.Frame(self.settings_frame)
        repeat_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.title_repeat_check = ttk.Checkbutton(repeat_frame, text="제목 반복 삽입", variable=self.title_repeat_var)
        self.title_repeat_check.grid(row=0, column=0, padx=(0, 10))
        
        ttk.Label(repeat_frame, text="반복 횟수:").grid(row=0, column=1, padx=(0, 5))
        repeat_entry = ttk.Entry(repeat_frame, textvariable=self.repeat_count_var, width=10)
        repeat_entry.grid(row=0, column=2, padx=(0, 10))
        
        self.random_repeat_check = ttk.Checkbutton(repeat_frame, text="랜덤 반복", variable=self.random_repeat_var)
        self.random_repeat_check.grid(row=0, column=3, padx=(0, 10))
        
        ttk.Label(repeat_frame, text="(입력된 최대 횟수까지 랜덤 삽입 후 공백 2줄 추가)", foreground="gray").grid(row=0, column=4)
        
        # 제목 뒤 단어 추가 설정
        append_word_frame = ttk.Frame(self.settings_frame)
        append_word_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.append_word_check = ttk.Checkbutton(append_word_frame, text="제목 뒤 단어 추가", variable=self.append_word_var)
        self.append_word_check.grid(row=0, column=0, padx=(0, 10))
        
        ttk.Label(append_word_frame, text="추가 단어:").grid(row=0, column=1, padx=(0, 5))
        self.append_word_entry = ttk.Entry(append_word_frame, textvariable=self.append_word_text, width=20)
        self.append_word_entry.grid(row=0, column=2, padx=(0, 10))
        
        # 태그 개수 및 공백 입력
        tag_options_frame = ttk.Frame(self.tag_frame)
        tag_options_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Label(tag_options_frame, text="삽입 태그 개수:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        self.tag_count_entry = ttk.Entry(tag_options_frame, textvariable=self.tag_count_var, width=10)
        self.tag_count_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        
        ttk.Label(tag_options_frame, text="태그 삽입 공백:").grid(row=0, column=2, padx=(0, 10), sticky=tk.W)
        self.tag_spacing_entry = ttk.Entry(tag_options_frame, textvariable=self.tag_spacing_var, width=10)
        self.tag_spacing_entry.grid(row=0, column=3, sticky=tk.W)
        
        # 해시태그 편집창
        hashtag_edit_frame = ttk.Frame(self.settings_frame)
        hashtag_edit_frame.grid(row=100, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(hashtag_edit_frame, text="직접 입력 태그 (한 줄에 하나씩):").grid(row=0, column=0, sticky=tk.W)
        self.hashtag_text = scrolledtext.ScrolledText(hashtag_edit_frame, height=4, width=80)
        self.hashtag_text.grid(row=1, column=0, pady=(5, 0), sticky=(tk.W, tk.E))
        
        # 실행 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=(0, 10))
        
        ttk.Button(button_frame, text="실행", command=self.execute_processing, style="Accent.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="종료", command=self.root.quit).pack(side=tk.LEFT)
        
        # 상태 표시
        self.status_var = tk.StringVar(value="준비됨")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        manuscript_frame.columnconfigure(0, weight=1)
        title_frame.columnconfigure(1, weight=1)
        replace_frame.columnconfigure(2, weight=1)
        self.tag_frame.columnconfigure(1, weight=1)
        filter_frame.columnconfigure(1, weight=1)
        save_path_frame.columnconfigure(0, weight=1)
        path_select_frame.columnconfigure(1, weight=1)
        relative_path_frame.columnconfigure(1, weight=1)
        self.settings_frame.columnconfigure(0, weight=1)
        tag_options_frame.columnconfigure(1, weight=1)
        tag_options_frame.columnconfigure(3, weight=1)

    def on_tag_load_mode_change(self):
        """태그 로딩 모드 변경 시 UI 업데이트"""
        self.update_tag_files()

    def load_batch_tag_files(self):
        """일괄 태그 파일 로딩"""
        files = filedialog.askopenfilenames(
            title="태그 파일들을 선택하세요 (최대 50개)",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if files:
            if len(files) > 50:
                messagebox.showwarning("경고", "최대 50개까지만 선택할 수 있습니다. 처음 50개만 로드됩니다.")
                files = files[:50]
            
            # 태그 파일 개수 업데이트
            self.tag_file_count_var.set(str(len(files)))
            self.update_tag_files()
            
            # 파일들을 순차적으로 로드
            for i, file_path in enumerate(files):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        self.tag_files[i]["tags"] = [line.strip() for line in f if line.strip()]
                    self.tag_files[i]["path"] = file_path
                    self.tag_files[i]["label"].config(text=os.path.basename(file_path))
                    self.tag_files[i]["preview"].config(state=tk.NORMAL)
                    self.tag_files[i]["preview"].delete(1.0, tk.END)
                    self.tag_files[i]["preview"].insert(tk.END, "\n".join(self.tag_files[i]["tags"][:5]))
                    self.tag_files[i]["preview"].config(state=tk.DISABLED)
                except Exception as e:
                    messagebox.showerror("오류", f"태그 파일 {i+1} 로드 중 오류: {str(e)}")
                    continue
            
            self.status_var.set(f"일괄 로드 완료: {len(files)}개 파일")

    def update_tag_files(self):
        """태그 파일 입력란과 설정을 동적으로 생성"""
        try:
            count = int(self.tag_file_count_var.get())
            if count < 1 or count > 50:
                raise ValueError
        except ValueError:
            messagebox.showerror("오류", "태그 파일 개수는 1~50 사이의 숫자여야 합니다.")
            self.tag_file_count_var.set("3")
            count = 3
        
        # 기존 위젯 제거
        for widget in self.tag_files_frame.winfo_children():
            widget.destroy()
        
        # 기존 설정 섹션의 태그 설정 제거
        if hasattr(self, 'settings_frame') and self.settings_frame.winfo_exists():
            for widget in self.settings_frame.winfo_children():
                if isinstance(widget, ttk.LabelFrame) and widget.cget("text").startswith("태그 파일"):
                    widget.destroy()
        
        # 새로운 태그 파일과 설정 초기화
        self.tag_files = [{"path": "", "tags": [], "label": None, "preview": None, "insert_mode": tk.StringVar(value="fixed")} for _ in range(count)]
        self.tag_settings = [
            {
                "hashtag_var": tk.BooleanVar(),
                "position": tk.StringVar(value="bottom"),
                "random_position": tk.BooleanVar()
            } for _ in range(count)
        ]
        
        # 태그 파일 입력 위젯 생성
        for i in range(count):
            row_frame = ttk.Frame(self.tag_files_frame)
            row_frame.grid(row=i, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
            
            # 개별 불러오기 모드일 때만 개별 버튼 표시
            if self.tag_load_mode_var.get() == "individual":
                ttk.Button(row_frame, text=f"태그 파일 {i+1} 불러오기", command=lambda x=i: self.load_tag_file(x)).grid(row=0, column=0, padx=(0, 10))
            else:
                ttk.Label(row_frame, text=f"태그 파일 {i+1}:").grid(row=0, column=0, padx=(0, 10))
            
            self.tag_files[i]["label"] = ttk.Label(row_frame, text=f"태그 파일 {i+1}이 선택되지 않았습니다.")
            self.tag_files[i]["label"].grid(row=0, column=1, sticky=tk.W)
            
            mode_frame = ttk.Frame(row_frame)
            mode_frame.grid(row=0, column=2, padx=(10, 0))
            ttk.Radiobutton(mode_frame, text="고정삽입", variable=self.tag_files[i]["insert_mode"], value="fixed").grid(row=0, column=0, padx=(0, 5))
            ttk.Radiobutton(mode_frame, text="설정값삽입", variable=self.tag_files[i]["insert_mode"], value="settings").grid(row=0, column=1)
            
            self.tag_files[i]["preview"] = scrolledtext.ScrolledText(self.tag_files_frame, height=2, width=80, state=tk.DISABLED)
            self.tag_files[i]["preview"].grid(row=i+count, column=0, columnspan=3, pady=(5, 5), sticky=(tk.W, tk.E))
        
        # 태그 설정 섹션 생성
        base_row = 3
        for i in range(count):
            tag_settings_frame = ttk.LabelFrame(self.settings_frame, text=f"태그 파일 {i+1} 설정", padding="10")
            tag_settings_frame.grid(row=base_row + i, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
            
            ttk.Checkbutton(tag_settings_frame, text="태그 삽입", variable=self.tag_settings[i]["hashtag_var"]).grid(row=0, column=0, padx=(0, 10))
            
            position_frame = ttk.Frame(tag_settings_frame)
            position_frame.grid(row=0, column=1, padx=(10, 0))
            
            ttk.Radiobutton(position_frame, text="상단", variable=self.tag_settings[i]["position"], value="top").grid(row=0, column=0, padx=(0, 5))
            ttk.Radiobutton(position_frame, text="중단", variable=self.tag_settings[i]["position"], value="middle").grid(row=0, column=1, padx=(0, 5))
            ttk.Radiobutton(position_frame, text="하단", variable=self.tag_settings[i]["position"], value="bottom").grid(row=0, column=2)
            
            ttk.Checkbutton(position_frame, text="랜덤 위치", variable=self.tag_settings[i]["random_position"]).grid(row=0, column=3, padx=(5, 0))
        
        self.tag_files_frame.columnconfigure(1, weight=1)

    def add_manuscript_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        for file in files:
            if file not in self.manuscript_files:
                self.manuscript_files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
        self.status_var.set(f"{len(self.manuscript_files)}개의 원고 파일이 추가됨")

    def remove_selected_file(self):
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.file_listbox.delete(index)
            self.manuscript_files.pop(index)
            self.status_var.set("선택된 파일이 제거됨")
        else:
            messagebox.showwarning("경고", "제거할 파일을 선택해주세요.")

    def clear_all_files(self):
        self.file_listbox.delete(0, tk.END)
        self.manuscript_files.clear()
        self.status_var.set("모든 파일이 삭제됨")

    def load_title_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.title_list = [line.strip() for line in f if line.strip()]
                self.title_file_path = file_path
                self.title_file_label.config(text=os.path.basename(file_path))
                self.title_preview.config(state=tk.NORMAL)
                self.title_preview.delete(1.0, tk.END)
                self.title_preview.insert(tk.END, "\n".join(self.title_list[:5]))
                self.title_preview.config(state=tk.DISABLED)
                self.status_var.set("제목 파일이 로드됨")
            except Exception as e:
                messagebox.showerror("오류", f"제목 파일 로드 중 오류: {str(e)}")

    def load_tag_file(self, index):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.tag_files[index]["tags"] = [line.strip() for line in f if line.strip()]
                self.tag_files[index]["path"] = file_path
                self.tag_files[index]["label"].config(text=os.path.basename(file_path))
                self.tag_files[index]["preview"].config(state=tk.NORMAL)
                self.tag_files[index]["preview"].delete(1.0, tk.END)
                self.tag_files[index]["preview"].insert(tk.END, "\n".join(self.tag_files[index]["tags"][:5]))
                self.tag_files[index]["preview"].config(state=tk.DISABLED)
                self.status_var.set(f"태그 파일 {index+1}이 로드됨")
            except Exception as e:
                messagebox.showerror("오류", f"태그 파일 {index+1} 로드 중 오류: {str(e)}")

    def load_filter_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.filter_words = [line.strip() for line in f if line.strip()]
                self.filter_file_path = file_path
                self.filter_file_label.config(text=os.path.basename(file_path))
                self.status_var.set("필터링 파일이 로드됨")
            except Exception as e:
                messagebox.showerror("오류", f"필터링 파일 로드 중 오류: {str(e)}")

    def select_save_path(self):
        path = filedialog.askdirectory()
        if path:
            self.save_path = path
            self.save_path_label.config(text=os.path.basename(path))
            self.status_var.set("저장 경로가 설정됨")

    def get_save_file_path(self, original_path, title_index=None):
        base_name = os.path.splitext(os.path.basename(original_path))[0]
        save_dir = ""
        
        if self.path_type_var.get() == "original":
            save_dir = os.path.dirname(original_path)
        elif self.path_type_var.get() == "absolute":
            save_dir = self.save_path
        elif self.path_type_var.get() == "relative":
            save_dir = os.path.join(os.path.dirname(original_path), self.relative_path_var.get())
        
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        # 제목 앞단어 추가 기능
        if self.add_title_word_to_filename_var.get() and title_index is not None and self.title_list:
            title = self.title_list[title_index]
            first_word = self.get_first_word_from_title(title)
            if first_word:
                base_name = f"{first_word}_{base_name}"
        
        return os.path.join(save_dir, f"{base_name}_processed.txt")

    def get_first_word_from_title(self, title):
        """제목에서 첫 번째 단어를 추출하는 함수"""
        return title.split()[0] if title.split() else ""

    def replace_text_with_count(self, content, target, replacement, count):
        """지정된 횟수만큼 문자열을 치환하는 함수"""
        if count <= 0:
            return content
        
        replaced_count = 0
        result = content
        
        while replaced_count < count and target in result:
            result = result.replace(target, replacement, 1)
            replaced_count += 1
        
        return result

    def execute_processing(self):
        """메인 실행 함수"""
        if not self.manuscript_files:
            messagebox.showwarning("경고", "원고 파일을 먼저 추가해주세요.")
            return
        
        if not self.title_list:
            messagebox.showwarning("경고", "제목 파일을 먼저 불러와주세요.")
            return
        
        if self.path_type_var.get() == "absolute" and not self.save_path:
            messagebox.showwarning("경고", "절대경로 저장을 선택했지만 저장 폴더가 설정되지 않았습니다.")
            return
        
        try:
            tag_count = int(self.tag_count_var.get())
            if tag_count < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("오류", "삽입 태그 개수는 1 이상의 숫자여야 합니다.")
            return
        
        try:
            tag_spacing = int(self.tag_spacing_var.get())
            if tag_spacing < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("오류", "태그 삽입 공백은 0 이상의 숫자여야 합니다.")
            return
        
        try:
            max_repeat_count = int(self.repeat_count_var.get())
            if max_repeat_count < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("오류", "반복 횟수는 1 이상의 숫자여야 합니다.")
            return
        
        # 치환 횟수 검증 추가
        try:
            replace_count = int(self.replace_count_var.get())
            if replace_count < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("오류", "치환 횟수는 1 이상의 숫자여야 합니다.")
            return
        
        # 제목 뒤 추가 단어 확인
        append_word = self.append_word_text.get().strip() if self.append_word_var.get() else ""
        
        # 치환 기능 설정 확인
        replace_enabled = self.replace_enabled_var.get()
        replace_target = self.replace_target_var.get().strip()
        
        success_count = 0
        
        for i, file_path in enumerate(self.manuscript_files):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    original_content = f.read()
                
                # 제목 인덱스 계산
                title_index = i % len(self.title_list)
                
                # 치환 기능 처리 (지정된 횟수만큼)
                content_to_process = original_content
                if replace_enabled and replace_target and self.title_list:
                    title = self.title_list[title_index]
                    first_word = self.get_first_word_from_title(title)
                    if first_word:
                        content_to_process = self.replace_text_with_count(original_content, replace_target, first_word, replace_count)
                
                # 제목 처리
                if self.title_mode_var.get() == "single":
                    title = self.title_list[title_index]
                    if append_word:
                        title += append_word
                    if self.random_repeat_var.get():
                        repeat_count = random.randint(1, max_repeat_count)
                    else:
                        repeat_count = max_repeat_count if self.title_repeat_var.get() else 1
                    title_section = (title + '\n') * repeat_count
                else:
                    start_index = (i * max_repeat_count) % len(self.title_list)
                    titles = []
                    for j in range(max_repeat_count):
                        current_title_index = (start_index + j) % len(self.title_list)
                        title = self.title_list[current_title_index]
                        if append_word:
                            title += append_word
                        titles.append(title)
                    title_section = '\n'.join(titles)
                
                title_section += '\n' * 2
                
                filtered_content = content_to_process
                if self.apply_filter_var.get() and self.filter_words:
                    for word in self.filter_words:
                        filtered_content = filtered_content.replace(word, "")
                
                # 태그 처리
                content_lines = (title_section + filtered_content).splitlines()
                new_content_lines = content_lines.copy()
                
                for idx in range(len(self.tag_files)):
                    if self.tag_files[idx]["tags"] and (self.tag_files[idx]["insert_mode"].get() == "fixed" or self.tag_settings[idx]["hashtag_var"].get()):
                        start_index = (i * tag_count) % len(self.tag_files[idx]["tags"])
                        selected_tags = []
                        for j in range(tag_count):
                            tag_index = (start_index + j) % len(self.tag_files[idx]["tags"])
                            selected_tags.append(self.tag_files[idx]["tags"][tag_index])
                        hashtag_section = '\n'.join(selected_tags) + '\n'
                        
                        # 삽입 위치 결정
                        if self.tag_files[idx]["insert_mode"].get() == "fixed":
                            position = "bottom"
                        else:
                            if self.tag_settings[idx]["random_position"].get():
                                position = random.choice(["top", "middle", "bottom"])
                            else:
                                position = self.tag_settings[idx]["position"].get()
                        
                        # 삽입 위치에 따라 태그 삽입
                        if position == "top":
                            new_content_lines.insert(0, hashtag_section.strip())
                        elif position == "middle":
                            mid_point = len(new_content_lines) // 2
                            new_content_lines.insert(mid_point, hashtag_section.strip())
                        elif position == "bottom":
                            new_content_lines.append('\n' * tag_spacing + hashtag_section.strip())
                
                new_content = '\n'.join(new_content_lines)
                
                save_file_path = self.get_save_file_path(file_path, title_index)
                
                with open(save_file_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                
                success_count += 1
                
            except Exception as e:
                messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다:\n{os.path.basename(file_path)}\n{str(e)}")
                continue
        
        if success_count > 0:
            messagebox.showinfo("완료", f"{success_count}개 파일이 성공적으로 처리되었습니다.")
            self.status_var.set(f"처리 완료: {success_count}개 파일")
        else:
            messagebox.showerror("오류", "처리된 파일이 없습니다.")

def main():
    root = tk.Tk()
    app = TitleInserterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()