import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import hashlib
from collections import defaultdict
from difflib import SequenceMatcher
import re
import threading

class TextDuplicateChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("텍스트 중복 확인 프로그램")
        self.root.geometry("900x750")
        
        self.files = []  # 파일 경로 저장
        self.file_contents = {}  # 파일별 원본 내용 저장
        self.is_processing = False  # 처리 중 플래그
        
        self.create_widgets()
    
    def create_widgets(self):
        # 상단 프레임 - 기법 선택
        top_frame = ttk.LabelFrame(self.root, text="중복 검사 기법 선택", padding=10)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        self.method_var = tk.StringVar(value="hashing")
        methods = [
            ("해싱 (Hashing) - 완전 일치", "hashing"),
            ("N-gram 분석 - 유사 패턴", "ngram"),
            ("레벤슈타인 거리 - 편집 거리", "levenshtein"),
            ("지문법 (Winnowing) - 대용량 샘플링", "winnowing")
        ]
        
        for i, (text, value) in enumerate(methods):
            ttk.Radiobutton(top_frame, text=text, variable=self.method_var, 
                           value=value).grid(row=0, column=i, padx=5, sticky="w")
        
        # 레벤슈타인 설정 프레임
        lev_frame = ttk.Frame(top_frame)
        lev_frame.grid(row=1, column=0, columnspan=4, pady=5, sticky="w")
        
        ttk.Label(lev_frame, text="레벤슈타인 임계값:").pack(side="left", padx=5)
        self.lev_threshold = tk.IntVar(value=5)
        ttk.Spinbox(lev_frame, from_=1, to=20, textvariable=self.lev_threshold, 
                   width=5).pack(side="left", padx=5)
        
        ttk.Label(lev_frame, text="최소 문장 길이:").pack(side="left", padx=15)
        self.min_length = tk.IntVar(value=10)
        ttk.Spinbox(lev_frame, from_=5, to=100, textvariable=self.min_length, 
                   width=5).pack(side="left", padx=5)
        
        # 중간 프레임 - 파일 관리
        middle_frame = ttk.LabelFrame(self.root, text="파일 관리", padding=10)
        middle_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 버튼 프레임
        button_frame = ttk.Frame(middle_frame)
        button_frame.pack(fill="x", pady=(0, 5))
        
        ttk.Button(button_frame, text="파일 추가", 
                  command=self.add_files).pack(side="left", padx=5)
        ttk.Button(button_frame, text="선택 파일 제거", 
                  command=self.remove_file).pack(side="left", padx=5)
        ttk.Button(button_frame, text="모두 제거", 
                  command=self.clear_files).pack(side="left", padx=5)
        
        self.check_button = ttk.Button(button_frame, text="중복 검사 실행", 
                  command=self.check_duplicates)
        self.check_button.pack(side="left", padx=20)
        
        self.cancel_button = ttk.Button(button_frame, text="검사 중지", 
                  command=self.cancel_check, state="disabled")
        self.cancel_button.pack(side="left", padx=5)
        
        # 파일 리스트뷰
        list_frame = ttk.Frame(middle_frame)
        list_frame.pack(fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.file_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, 
                                       height=8)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # 프로그레스 바
        progress_frame = ttk.Frame(self.root)
        progress_frame.pack(fill="x", padx=10, pady=5)
        
        self.progress_label = ttk.Label(progress_frame, text="대기 중...")
        self.progress_label.pack(side="left", padx=5)
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress.pack(side="left", fill="x", expand=True, padx=5)
        
        # 하단 프레임 - 결과 표시
        bottom_frame = ttk.LabelFrame(self.root, text="검사 결과", padding=10)
        bottom_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.result_text = scrolledtext.ScrolledText(bottom_frame, height=15, 
                                                     wrap=tk.WORD)
        self.result_text.pack(fill="both", expand=True)
        
        # 태그 설정 (하이라이트용)
        self.result_text.tag_config("highlight", background="yellow", 
                                   foreground="red", font=("", 10, "bold"))
        self.result_text.tag_config("header", foreground="blue", 
                                   font=("", 10, "bold"))
        self.result_text.tag_config("warning", foreground="orange")
    
    def add_files(self):
        """파일 추가"""
        files = filedialog.askopenfilenames(
            title="TXT 파일 선택",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        for file in files:
            if file not in self.files:
                self.files.append(file)
                self.file_listbox.insert(tk.END, file.split("/")[-1])
                
                # 파일 내용 읽기 (원본 보존)
                try:
                    with open(file, 'r', encoding='utf-8') as f:
                        self.file_contents[file] = f.read()
                except:
                    try:
                        with open(file, 'r', encoding='cp949') as f:
                            self.file_contents[file] = f.read()
                    except Exception as e:
                        messagebox.showerror("오류", f"파일 읽기 실패: {file}\n{e}")
    
    def remove_file(self):
        """선택된 파일 제거"""
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            file_path = self.files[index]
            del self.file_contents[file_path]
            del self.files[index]
            self.file_listbox.delete(index)
    
    def clear_files(self):
        """모든 파일 제거"""
        self.files.clear()
        self.file_contents.clear()
        self.file_listbox.delete(0, tk.END)
        self.result_text.delete(1.0, tk.END)
    
    def preprocess_text(self, text):
        """텍스트 전처리"""
        # 공백 정리 (여러 공백을 하나로)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def cancel_check(self):
        """검사 중지"""
        self.is_processing = False
        self.progress_label.config(text="검사가 중지되었습니다.")
    
    def check_duplicates(self):
        """중복 검사 실행"""
        if not self.files:
            messagebox.showwarning("경고", "파일을 추가해주세요.")
            return
        
        if self.is_processing:
            messagebox.showwarning("경고", "이미 검사가 진행 중입니다.")
            return
        
        self.result_text.delete(1.0, tk.END)
        method = self.method_var.get()
        
        self.result_text.insert(tk.END, f"=== 검사 기법: {method.upper()} ===\n\n", "header")
        
        # 쓰레드로 실행
        self.is_processing = True
        self.check_button.config(state="disabled")
        self.cancel_button.config(state="normal")
        self.progress['value'] = 0
        
        thread = threading.Thread(target=self.run_check, args=(method,), daemon=True)
        thread.start()
    
    def run_check(self, method):
        """백그라운드에서 검사 실행"""
        try:
            if method == "hashing":
                self.check_with_hashing()
            elif method == "ngram":
                self.check_with_ngram()
            elif method == "levenshtein":
                self.check_with_levenshtein()
            elif method == "winnowing":
                self.check_with_winnowing()
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("오류", f"검사 중 오류 발생:\n{e}"))
        finally:
            self.root.after(0, self.finish_check)
    
    def finish_check(self):
        """검사 완료 후 처리"""
        self.is_processing = False
        self.check_button.config(state="normal")
        self.cancel_button.config(state="disabled")
        self.progress['value'] = 100
        if self.progress_label.cget("text") != "검사가 중지되었습니다.":
            self.progress_label.config(text="검사 완료!")
    
    def update_progress(self, current, total, message=""):
        """진행률 업데이트"""
        if total > 0:
            progress_value = (current / total) * 100
            self.root.after(0, lambda: self.progress.config(value=progress_value))
        if message:
            self.root.after(0, lambda: self.progress_label.config(text=message))
    
    def append_result(self, text, tag=None):
        """결과 텍스트 추가 (쓰레드 안전)"""
        def insert():
            if tag:
                self.result_text.insert(tk.END, text, tag)
            else:
                self.result_text.insert(tk.END, text)
            self.result_text.see(tk.END)
        self.root.after(0, insert)
    
    def check_with_hashing(self):
        """해싱 기법으로 중복 검사"""
        hash_map = defaultdict(list)
        
        total_files = len(self.files)
        for file_idx, file_path in enumerate(self.files):
            if not self.is_processing:
                return
            
            self.update_progress(file_idx, total_files, 
                               f"파일 분석 중... ({file_idx+1}/{total_files})")
            
            content = self.file_contents[file_path]
            preprocessed = self.preprocess_text(content)
            
            # 줄 단위로 분리
            lines = preprocessed.split('\n')
            
            for i, line in enumerate(lines):
                if line.strip():
                    # 해시값 생성
                    hash_value = hashlib.md5(line.encode()).hexdigest()
                    hash_map[hash_value].append({
                        'file': file_path.split("/")[-1],
                        'line_num': i + 1,
                        'content': line
                    })
        
        # 중복 결과 출력
        self.update_progress(0, 1, "결과 생성 중...")
        duplicate_count = 0
        for hash_value, occurrences in hash_map.items():
            if not self.is_processing:
                return
            
            if len(occurrences) > 1:
                duplicate_count += 1
                self.append_result(f"[중복 #{duplicate_count}] ", "highlight")
                self.append_result(f"총 {len(occurrences)}회 발견\n")
                self.append_result(f"내용: {occurrences[0]['content']}\n")
                self.append_result("위치:\n")
                
                for occ in occurrences:
                    self.append_result(f"  - {occ['file']} (줄 {occ['line_num']})\n")
                self.append_result("\n")
        
        if duplicate_count == 0:
            self.append_result("중복된 내용이 없습니다.\n")
        else:
            self.append_result(f"\n총 {duplicate_count}개의 중복 항목을 발견했습니다.\n", "header")
    
    def check_with_ngram(self):
        """N-gram 분석"""
        n = 3  # 3-gram 사용
        ngram_map = defaultdict(list)
        
        total_files = len(self.files)
        for file_idx, file_path in enumerate(self.files):
            if not self.is_processing:
                return
            
            self.update_progress(file_idx, total_files, 
                               f"N-gram 생성 중... ({file_idx+1}/{total_files})")
            
            content = self.file_contents[file_path]
            preprocessed = self.preprocess_text(content)
            
            # 단어 단위로 분리
            words = preprocessed.split()
            
            # N-gram 생성
            for i in range(len(words) - n + 1):
                ngram = ' '.join(words[i:i+n])
                ngram_map[ngram].append({
                    'file': file_path.split("/")[-1],
                    'position': i,
                    'content': ngram
                })
        
        # 중복 결과 출력
        self.update_progress(0, 1, "결과 생성 중...")
        duplicate_count = 0
        for ngram, occurrences in ngram_map.items():
            if not self.is_processing:
                return
            
            if len(occurrences) > 1:
                duplicate_count += 1
                self.append_result(f"[중복 패턴 #{duplicate_count}] ", "highlight")
                self.append_result(f"총 {len(occurrences)}회 발견\n")
                self.append_result(f"패턴: {ngram}\n")
                self.append_result("위치:\n")
                
                for occ in occurrences:
                    self.append_result(f"  - {occ['file']} (단어 위치 {occ['position']})\n")
                self.append_result("\n")
        
        if duplicate_count == 0:
            self.append_result("중복된 패턴이 없습니다.\n")
        else:
            self.append_result(f"\n총 {duplicate_count}개의 중복 패턴을 발견했습니다.\n", "header")
    
    def levenshtein_distance(self, s1, s2):
        """레벤슈타인 거리 계산 (최적화 버전)"""
        if len(s1) < len(s2):
            return self.levenshtein_distance(s2, s1)
        
        if len(s2) == 0:
            return len(s1)
        
        # 메모리 최적화: 이전 행만 저장
        previous_row = range(len(s2) + 1)
        for i, c1 in enumerate(s1):
            current_row = [i + 1]
            for j, c2 in enumerate(s2):
                insertions = previous_row[j + 1] + 1
                deletions = current_row[j] + 1
                substitutions = previous_row[j] + (c1 != c2)
                current_row.append(min(insertions, deletions, substitutions))
            previous_row = current_row
        
        return previous_row[-1]
    
    def check_with_levenshtein(self):
        """레벤슈타인 거리로 유사 문장 찾기 (최적화 버전)"""
        threshold = self.lev_threshold.get()
        min_len = self.min_length.get()
        
        # 모든 줄 수집
        all_lines = []
        for file_path in self.files:
            if not self.is_processing:
                return
            
            content = self.file_contents[file_path]
            preprocessed = self.preprocess_text(content)
            lines = preprocessed.split('\n')
            
            for i, line in enumerate(lines):
                line_stripped = line.strip()
                # 최소 길이 필터링
                if len(line_stripped) >= min_len:
                    all_lines.append({
                        'file': file_path.split("/")[-1],
                        'line_num': i + 1,
                        'content': line_stripped
                    })
        
        total_lines = len(all_lines)
        
        # 너무 많은 줄이면 경고
        if total_lines > 1000:
            self.append_result(
                f"⚠️ 경고: {total_lines}개의 줄을 비교합니다. 시간이 오래 걸릴 수 있습니다.\n", 
                "warning")
        
        # 모든 줄 쌍 비교
        found_pairs = []
        total_comparisons = (total_lines * (total_lines - 1)) // 2
        comparison_count = 0
        
        for i in range(len(all_lines)):
            if not self.is_processing:
                return
            
            # 진행률 업데이트 (100번마다)
            if i % 10 == 0:
                progress = (i / total_lines) * 100
                self.update_progress(i, total_lines, 
                    f"비교 중... {i}/{total_lines} 줄 ({len(found_pairs)}개 발견)")
            
            for j in range(i + 1, len(all_lines)):
                comparison_count += 1
                
                line1 = all_lines[i]['content']
                line2 = all_lines[j]['content']
                
                # 길이 차이가 임계값보다 크면 스킵 (최적화)
                if abs(len(line1) - len(line2)) > threshold:
                    continue
                
                distance = self.levenshtein_distance(line1, line2)
                
                if 0 < distance <= threshold:
                    found_pairs.append({
                        'distance': distance,
                        'line1': all_lines[i],
                        'line2': all_lines[j]
                    })
        
        # 결과 출력
        self.update_progress(0, 1, "결과 생성 중...")
        
        if not found_pairs:
            self.append_result(f"유사한 문장이 없습니다. (임계값: {threshold})\n")
        else:
            found_pairs.sort(key=lambda x: x['distance'])
            
            # 최대 100개만 표시
            display_count = min(len(found_pairs), 100)
            if len(found_pairs) > 100:
                self.append_result(
                    f"⚠️ {len(found_pairs)}개 발견, 상위 100개만 표시합니다.\n\n", 
                    "warning")
            
            for idx, pair in enumerate(found_pairs[:display_count], 1):
                if not self.is_processing:
                    return
                
                self.append_result(f"[유사 #{idx}] ", "highlight")
                self.append_result(f"편집 거리: {pair['distance']}\n")
                self.append_result(f"문장1: {pair['line1']['content']}\n")
                self.append_result(
                    f"  위치: {pair['line1']['file']} (줄 {pair['line1']['line_num']})\n")
                self.append_result(f"문장2: {pair['line2']['content']}\n")
                self.append_result(
                    f"  위치: {pair['line2']['file']} (줄 {pair['line2']['line_num']})\n\n")
            
            self.append_result(
                f"\n총 {len(found_pairs)}개의 유사 문장 쌍을 발견했습니다.\n", "header")
    
    def check_with_winnowing(self):
        """Winnowing 기법 (지문법)"""
        k = 5  # k-gram 크기
        window_size = 4  # 윈도우 크기
        
        def get_kgrams(text):
            """k-gram 생성"""
            return [text[i:i+k] for i in range(len(text) - k + 1)]
        
        def hash_kgrams(kgrams):
            """k-gram 해싱"""
            return [hash(kg) for kg in kgrams]
        
        def winnow(hashes):
            """Winnowing 알고리즘"""
            fingerprints = []
            for i in range(len(hashes) - window_size + 1):
                window = hashes[i:i+window_size]
                min_hash = min(window)
                min_idx = i + window.index(min_hash)
                fingerprints.append((min_idx, min_hash))
            
            # 중복 제거
            seen = set()
            unique_fp = []
            for fp in fingerprints:
                if fp not in seen:
                    seen.add(fp)
                    unique_fp.append(fp)
            return unique_fp
        
        # 각 파일의 지문 생성
        file_fingerprints = {}
        total_files = len(self.files)
        
        for file_idx, file_path in enumerate(self.files):
            if not self.is_processing:
                return
            
            self.update_progress(file_idx, total_files, 
                               f"지문 생성 중... ({file_idx+1}/{total_files})")
            
            content = self.file_contents[file_path]
            preprocessed = self.preprocess_text(content)
            
            kgrams = get_kgrams(preprocessed)
            hashes = hash_kgrams(kgrams)
            fingerprints = winnow(hashes)
            
            file_fingerprints[file_path.split("/")[-1]] = {
                'fingerprints': fingerprints,
                'content': preprocessed
            }
        
        # 지문 비교로 중복 찾기
        self.update_progress(0, 1, "중복 검색 중...")
        fingerprint_map = defaultdict(list)
        for filename, data in file_fingerprints.items():
            for pos, fp in data['fingerprints']:
                fingerprint_map[fp].append({
                    'file': filename,
                    'position': pos
                })
        
        # 결과 출력
        duplicate_count = 0
        for fp, occurrences in fingerprint_map.items():
            if not self.is_processing:
                return
            
            if len(occurrences) > 1:
                duplicate_count += 1
                self.append_result(f"[중복 지문 #{duplicate_count}] ", "highlight")
                self.append_result(f"총 {len(occurrences)}회 발견\n")
                self.append_result("위치:\n")
                
                for occ in occurrences:
                    self.append_result(f"  - {occ['file']} (위치 {occ['position']})\n")
                self.append_result("\n")
        
        if duplicate_count == 0:
            self.append_result("중복된 지문이 없습니다.\n")
        else:
            self.append_result(
                f"\n총 {duplicate_count}개의 중복 지문을 발견했습니다.\n", "header")


if __name__ == "__main__":
    root = tk.Tk()
    app = TextDuplicateChecker(root)
    root.mainloop()