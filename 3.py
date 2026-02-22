import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
import pandas as pd
import os
import sys
import time
from datetime import datetime
import asyncio
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
import re
import csv
import urllib.parse

class NaverBlogRankChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("네이버 블로그 순위 체커")
        self.root.geometry("900x900")
        self.root.resizable(True, True)
        
        # 상태 변수
        self.is_running = False
        self.should_stop = False
        self.keywords = []
        self.blog_ids = []
        self.results_queue = queue.Queue()
        self.log_queue = queue.Queue()
        self.results = []
        self.rank_results = []
        self.last_save_time = None
        self.search_mode_var = tk.StringVar(value="relevance")
        
        # GUI 구성
        self.create_gui()
        
        # 로그 업데이트 스레드
        self.start_log_updater()
    
    def create_gui(self):
        # 메인 노트북 (탭)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 순위 체크 탭
        rank_frame = ttk.Frame(notebook)
        notebook.add(rank_frame, text="순위 체크")
        self.create_rank_tab(rank_frame)
        
        # 크롤링 탭
        crawl_frame = ttk.Frame(notebook)
        notebook.add(crawl_frame, text="블로그 크롤링")
        self.create_crawl_tab(crawl_frame)
        
        # 공통 컨트롤 프레임
        control_frame = ttk.Frame(self.root)
        control_frame.pack(padx=10, pady=10, fill="x")
        
        self.start_button = ttk.Button(control_frame, text="시작", command=self.start_processing)
        self.start_button.pack(side="left", padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="중지", command=self.stop_processing, state=tk.DISABLED)
        self.stop_button.pack(side="left", padx=5)
        
        ttk.Button(control_frame, text="종료", command=self.root.destroy).pack(side="right", padx=5)
    
    def create_rank_tab(self, parent):
        # 순위 체크 설정 프레임
        rank_settings_frame = ttk.LabelFrame(parent, text="순위 체크 설정")
        rank_settings_frame.pack(padx=10, pady=5, fill="x")
        
        # 키워드 파일 설정
        keyword_frame = ttk.Frame(rank_settings_frame)
        keyword_frame.pack(padx=5, pady=5, fill="x")
        
        ttk.Label(keyword_frame, text="키워드 파일:").pack(side="left", padx=5)
        self.keyword_path_var = tk.StringVar()
        ttk.Entry(keyword_frame, textvariable=self.keyword_path_var, width=50).pack(side="left", padx=5)
        ttk.Button(keyword_frame, text="찾기", command=self.load_keyword_file).pack(side="left", padx=5)
        
        # 블로그 ID 파일 설정
        id_frame = ttk.Frame(rank_settings_frame)
        id_frame.pack(padx=5, pady=5, fill="x")
        
        ttk.Label(id_frame, text="블로그 ID 파일:").pack(side="left", padx=5)
        self.id_path_var = tk.StringVar()
        ttk.Entry(id_frame, textvariable=self.id_path_var, width=50).pack(side="left", padx=5)
        ttk.Button(id_frame, text="찾기", command=self.load_id_file).pack(side="left", padx=5)
        
        # 검색 정렬 옵션
        search_mode_frame = ttk.Frame(rank_settings_frame)
        search_mode_frame.pack(padx=5, pady=5, fill="x")
        ttk.Label(search_mode_frame, text="검색 정렬:").pack(side="left", padx=5)
        ttk.Radiobutton(search_mode_frame, text="관련도순", variable=self.search_mode_var, value="relevance").pack(side="left", padx=5)
        ttk.Radiobutton(search_mode_frame, text="최신순", variable=self.search_mode_var, value="recency").pack(side="left", padx=5)
        
        # 저장 옵션 추가
        save_options_frame = ttk.LabelFrame(rank_settings_frame, text="저장 옵션")
        save_options_frame.pack(padx=5, pady=5, fill="x")
        
        self.save_option_var = tk.StringVar(value="matched_only")
        ttk.Radiobutton(save_options_frame, text="블로그ID만 저장 (매칭된 결과만)", 
                       variable=self.save_option_var, value="matched_only").pack(side="left", padx=5)
        ttk.Radiobutton(save_options_frame, text="모든 순위 저장 (1-10위 전체)", 
                       variable=self.save_option_var, value="all_ranks").pack(side="left", padx=5)
        
        # 동시 실행 갯수 설정
        concurrent_frame = ttk.Frame(rank_settings_frame)
        concurrent_frame.pack(padx=5, pady=5, fill="x")
        
        ttk.Label(concurrent_frame, text="동시 실행 갯수:").pack(side="left", padx=5)
        self.concurrent_count_var = tk.IntVar(value=3)
        ttk.Spinbox(concurrent_frame, from_=1, to=10, textvariable=self.concurrent_count_var, width=5).pack(side="left", padx=5)
        
        # 지연 시간 설정
        delay_frame = ttk.Frame(rank_settings_frame)
        delay_frame.pack(padx=5, pady=5, fill="x")
        
        ttk.Label(delay_frame, text="지연 시간(초):").pack(side="left", padx=5)
        self.delay_var = tk.DoubleVar(value=1.0)
        ttk.Spinbox(delay_frame, from_=0.1, to=10.0, increment=0.1, textvariable=self.delay_var, width=5).pack(side="left", padx=5)
    
    def create_crawl_tab(self, parent):
        # 크롤링 설정 프레임
        crawl_settings_frame = ttk.LabelFrame(parent, text="크롤링 설정")
        crawl_settings_frame.pack(padx=10, pady=5, fill="x")
        
        # 크롤링 모드 선택
        mode_frame = ttk.Frame(crawl_settings_frame)
        mode_frame.pack(padx=5, pady=5, fill="x")
        
        ttk.Label(mode_frame, text="크롤링 모드:").pack(side="left", padx=5)
        self.crawl_mode_var = tk.StringVar(value="disabled")
        ttk.Radiobutton(mode_frame, text="비활성화", variable=self.crawl_mode_var, value="disabled").pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="상세 정보 추출", variable=self.crawl_mode_var, value="enabled").pack(side="left", padx=5)
        
        # 추출할 정보 선택
        extract_info_frame = ttk.Frame(crawl_settings_frame)
        extract_info_frame.pack(padx=5, pady=5, fill="x")
        
        ttk.Label(extract_info_frame, text="추출할 정보:").pack(side="left", padx=5)
        
        self.extract_title_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(extract_info_frame, text="제목", variable=self.extract_title_var).pack(side="left", padx=5)
        
        self.extract_content_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(extract_info_frame, text="내용", variable=self.extract_content_var).pack(side="left", padx=5)
        
        self.extract_contact_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(extract_info_frame, text="연락처", variable=self.extract_contact_var).pack(side="left", padx=5)
        
        self.extract_talktalk_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(extract_info_frame, text="톡톡 링크", variable=self.extract_talktalk_var).pack(side="left", padx=5)
        
        # 저장 형식 선택
        save_format_frame = ttk.Frame(crawl_settings_frame)
        save_format_frame.pack(padx=5, pady=5, fill="x")
        
        ttk.Label(save_format_frame, text="저장 형식:").pack(side="left", padx=5)
        
        self.save_format_var = tk.StringVar(value="excel")
        ttk.Radiobutton(save_format_frame, text="엑셀(.xlsx)", variable=self.save_format_var, value="excel").pack(side="left", padx=5)
        ttk.Radiobutton(save_format_frame, text="텍스트(.txt)", variable=self.save_format_var, value="txt").pack(side="left", padx=5)
        ttk.Radiobutton(save_format_frame, text="CSV(.csv)", variable=self.save_format_var, value="csv").pack(side="left", padx=5)
        
        # 상태 및 로그 창 (공통)
        status_frame = ttk.LabelFrame(parent, text="상태 및 로그")
        status_frame.pack(padx=10, pady=5, fill="both", expand=True)
        
        self.log_text = scrolledtext.ScrolledText(status_frame, wrap=tk.WORD, height=15)
        self.log_text.pack(padx=5, pady=5, fill="both", expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 진행 상황 표시
        progress_frame = ttk.Frame(status_frame)
        progress_frame.pack(padx=5, pady=5, fill="x")
        
        ttk.Label(progress_frame, text="진행 상태:").pack(side="left", padx=5)
        self.progress_var = tk.StringVar(value="대기 중")
        ttk.Label(progress_frame, textvariable=self.progress_var).pack(side="left", padx=5)
        
        self.progress_bar = ttk.Progressbar(status_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(padx=5, pady=5, fill="x")
    
    def load_keyword_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("텍스트 파일", "*.txt")])
        if file_path:
            self.keyword_path_var.set(file_path)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.keywords = [line.strip() for line in f if line.strip()]
                self.log(f"키워드 파일 로드 완료: {len(self.keywords)}개의 키워드")
            except Exception as e:
                self.log(f"키워드 파일 로드 실패: {str(e)}")
                self.keywords = []
    
    def load_id_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("텍스트 파일", "*.txt")])
        if file_path:
            self.id_path_var.set(file_path)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.blog_ids = [line.strip().lower() for line in f if line.strip()]
                self.log(f"블로그 ID 파일 로드 완료: {len(self.blog_ids)}개의 ID")
            except Exception as e:
                self.log(f"블로그 ID 파일 로드 실패: {str(e)}")
                self.blog_ids = []
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        self.log_queue.put(log_message)
    
    def start_log_updater(self):
        def update_log():
            while True:
                try:
                    log_message = self.log_queue.get(block=False)
                    self.log_text.config(state=tk.NORMAL)
                    self.log_text.insert(tk.END, log_message + "\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                except queue.Empty:
                    break
            self.root.after(100, update_log)
        
        self.root.after(100, update_log)
    
    def start_processing(self):
        if not self.keywords:
            messagebox.showwarning("경고", "키워드 파일을 먼저 로드해주세요.")
            return
        
        if not self.blog_ids:
            messagebox.showwarning("경고", "블로그 ID 파일을 먼저 로드해주세요.")
            return
        
        self.is_running = True
        self.should_stop = False
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_var.set("처리 준비 중...")
        self.progress_bar["value"] = 0
        self.results = []
        self.rank_results = []
        
        # 처리 스레드 시작
        threading.Thread(target=self.processing_thread, daemon=True).start()
    
    def stop_processing(self):
        self.should_stop = True
        self.log("중지 요청 처리 중... 현재 작업이 완료될 때까지 기다려주세요.")
        self.stop_button.config(state=tk.DISABLED)
    
    def processing_thread(self):
        try:
            # 저장 경로 설정
            save_dir = filedialog.askdirectory(title="결과를 저장할 폴더를 선택하세요")
            if not save_dir:
                self.log("저장 경로가 선택되지 않았습니다. 처리를 중단합니다.")
                self.reset_ui()
                return
            
            # asyncio 이벤트 루프 실행
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            loop.run_until_complete(self.process_all_keywords(save_dir))
            
            # 최종 결과 저장
            if self.rank_results:
                self.save_rank_results(self.rank_results, save_dir)
            if self.results and self.crawl_mode_var.get() == "enabled":
                self.save_crawl_results(self.results, save_dir)
            
            self.log("모든 작업 완료")
            messagebox.showinfo("완료", "작업이 완료되었습니다.")
        except Exception as e:
            self.log(f"처리 중 오류 발생: {str(e)}")
        finally:
            self.reset_ui()
    
    async def process_all_keywords(self, save_dir):
        """모든 키워드를 처리하는 메인 함수"""
        self.log(f"총 {len(self.keywords)} 개의 키워드에 대한 순위 체크를 시작합니다.")
        
        # 작업 큐 생성
        task_queue = asyncio.Queue()
        for keyword in self.keywords:
            await task_queue.put(keyword)
        
        # 동시 실행 횟수
        concurrent_count = self.concurrent_count_var.get()
        
        # Playwright 초기화
        async with async_playwright() as p:
            # 작업자 태스크 생성
            workers = []
            for i in range(concurrent_count):
                task = asyncio.create_task(self.worker(p, task_queue, save_dir, i))
                workers.append(task)
            
            # 모든 작업이 완료될 때까지 대기
            await asyncio.gather(*workers)
    
    async def worker(self, playwright, task_queue, save_dir, worker_id):
        """개별 작업자 함수"""
        browser = await playwright.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        )
        
        page = await context.new_page()
        
        try:
            while not task_queue.empty() and not self.should_stop:
                try:
                    keyword = await task_queue.get()
                    self.log(f"워커 {worker_id}: '{keyword}' 처리 시작")
                    
                    # 순위 체크 수행
                    rank_result = await self.check_keyword_rank(page, keyword, worker_id)
                    if rank_result:
                        self.rank_results.extend(rank_result)
                    
                    # 크롤링 모드가 활성화된 경우 상세 정보 추출
                    if self.crawl_mode_var.get() == "enabled":
                        crawl_results = await self.crawl_keyword_posts(page, keyword, worker_id)
                        if crawl_results:
                            self.results.extend(crawl_results)
                    
                    total_items = len(self.keywords)
                    self.update_progress(total_items - task_queue.qsize(), total_items)
                    
                    # 요청 간 지연
                    await asyncio.sleep(self.delay_var.get())
                    
                except Exception as e:
                    self.log(f"키워드 '{keyword}' 처리 중 오류 발생: {str(e)}")
                finally:
                    task_queue.task_done()
        finally:
            await browser.close()
    
    def extract_blog_id(self, url):
        """URL에서 블로그 ID를 추출하는 개선된 함수"""
        if not url:
            return None
            
        try:
            # URL 디코딩
            decoded_url = urllib.parse.unquote(url)
            
            # URL 파싱
            parsed = urllib.parse.urlparse(decoded_url)
            path_parts = parsed.path.strip('/').split('/')
            query_params = urllib.parse.parse_qs(parsed.query)
            
            # 가능한 블로그 ID 추출 방법들
            candidates = []
            
            # 쿼리 파라미터에서 blogId 추출
            if 'blogId' in query_params:
                candidates.extend(query_params['blogId'])
            
            # 경로에서 추출: blog.naver.com/아이디/...
            if len(path_parts) >= 1 and 'blog.naver.com' in parsed.netloc:
                if path_parts[0] not in ['PostView.naver', 'PostView.nhn', '']:
                    candidates.append(path_parts[0])
            
            # PostView 경로인 경우
            if any(p in path_parts for p in ['PostView.naver', 'PostView.nhn']):
                if 'blogId' in query_params:
                    candidates.extend(query_params['blogId'])
            
            # 리다이렉트나 인코딩된 URL 처리
            if '%2F' in decoded_url:
                inner_url = urllib.parse.unquote(decoded_url.split('%2F', 1)[1] if '%2F' in decoded_url else decoded_url)
                inner_parsed = urllib.parse.urlparse(inner_url)
                inner_path = inner_parsed.path.strip('/').split('/')
                if inner_path:
                    candidates.append(inner_path[0])
            
            # 추가 패턴들
            profile_match = re.search(r'profile\?blogId=([^&]+)', decoded_url)
            if profile_match:
                candidates.append(profile_match.group(1))
            
            id_match = re.search(r'(?<=blog\.naver\.com/)[^/]+', decoded_url)
            if id_match:
                candidates.append(id_match.group(0))
            
            # 후보들 정제 및 유효성 검사
            valid_ids = []
            invalid_values = [
                'postview', 'postlist', 'blogmain', 'search', 
                'blog', 'naver', 'com', 'www', 'http', 'https',
                'post', 'view', 'main', 'list', 'home', 'index',
                'api', 'static', 'css', 'js', 'img', 'image'
            ]
            
            for cand in set(candidates):  # 중복 제거
                blog_id = cand.lower().strip()
                if (3 <= len(blog_id) <= 30 and  
                    blog_id not in invalid_values and
                    not blog_id.isdigit() and
                    re.match(r'^[a-z0-9_-]+$', blog_id)):
                    valid_ids.append(blog_id)
            
            if valid_ids:
                blog_id = valid_ids[0]
                return blog_id
                
        except Exception as e:
            self.log(f"블로그 ID 추출 오류: {url} - {str(e)}")
            
        return None
    
    async def check_keyword_rank(self, page, keyword, worker_id):
        """키워드 순위 체크 수행 - 안정성 개선"""
        try:
            # 키워드 URL 인코딩
            encoded_keyword = urllib.parse.quote(keyword)
            if self.search_mode_var.get() == "relevance":
                url = f"https://search.naver.com/search.naver?ssc=tab.blog.all&sm=tab_jum&query={encoded_keyword}"
            else:
                url = f"https://search.naver.com/search.naver?ssc=tab.blog.all&query={encoded_keyword}&sm=tab_opt&nso=so%3Add%2Cp%3Aall"
            
            self.log(f"워커 {worker_id}: '{keyword}' 검색 URL: {url}")
            await page.goto(url, timeout=60000)
            await page.wait_for_load_state('networkidle', timeout=30000)
            await asyncio.sleep(self.delay_var.get())
            
            # 페이지 스크롤
            try:
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                await asyncio.sleep(1)
            except:
                pass
            
            # 순위 정보 초기화
            rank_data = []
            
            # 더 안정적인 블로그 포스트 추출
            try:
                valid_posts = await self.extract_blog_posts_stable(page, worker_id, keyword)
            except Exception as e:
                self.log(f"워커 {worker_id}: '{keyword}' - 포스트 추출 중 오류: {str(e)}")
                # 기본 방식으로 폴백
                valid_posts = await self.extract_posts_fallback(page, worker_id, keyword)
            
            if not valid_posts:
                self.log(f"워커 {worker_id}: '{keyword}' - 유효한 블로그 포스트를 찾을 수 없습니다.")
                # 빈 순위 데이터라도 반환
                for rank in range(1, 11):
                    if self.save_option_var.get() == "all_ranks":
                        rank_data.append({
                            "키워드": keyword,
                            "순위": rank,
                            "상태": "NO",
                            "블로그ID": "",
                            "제목": "포스트 없음",
                            "URL": "",
                            "검색URL": url
                        })
                return rank_data
            
            self.log(f"워커 {worker_id}: '{keyword}' - 총 {len(valid_posts)}개의 유효한 블로그 포스트 발견")
            
            # 정확한 순위 체크 (1위부터 10위까지)
            for rank in range(1, 11):
                if rank - 1 < len(valid_posts):
                    post_info = valid_posts[rank - 1]
                    
                    status = "NO"
                    blog_id = ""
                    
                    if post_info.get('url'):
                        extracted_id = self.extract_blog_id(post_info['url'])
                        if extracted_id:
                            blog_id = extracted_id
                            if extracted_id in self.blog_ids:
                                status = "OK"
                                self.log(f"워커 {worker_id}: '{keyword}' - {rank}위에서 매칭 발견: {blog_id}")
                            else:
                                self.log(f"워커 {worker_id}: '{keyword}' - {rank}위 블로그 ID: {blog_id} (매칭되지 않음)")
                    
                    rank_data.append({
                        "키워드": keyword,
                        "순위": rank,
                        "상태": status,
                        "블로그ID": blog_id,
                        "제목": post_info.get('title', '제목 없음'),
                        "URL": post_info.get('url', ''),
                        "검색URL": url
                    })
                    
                else:
                    # 해당 순위에 포스트가 없는 경우 (모든 순위 저장 모드일 때만)
                    if self.save_option_var.get() == "all_ranks":
                        rank_data.append({
                            "키워드": keyword,
                            "순위": rank,
                            "상태": "NO",
                            "블로그ID": "",
                            "제목": "포스트 없음",
                            "URL": "",
                            "검색URL": url
                        })
            
            # 매칭된 결과 요약 로그
            matched_count = sum(1 for item in rank_data if item["상태"] == "OK")
            self.log(f"워커 {worker_id}: '{keyword}' 순위 체크 완료 - {matched_count}개 매칭 (총 {len(rank_data)}개 순위)")
            
            return rank_data
            
        except Exception as e:
            self.log(f"워커 {worker_id}: '{keyword}' 순위 체크 중 오류: {str(e)}")
            return []
    
    async def extract_blog_posts_stable(self, page, worker_id, keyword):
        """안정적인 블로그 포스트 추출 방법"""
        try:
            valid_posts = []
            
            # 1단계: 모든 블로그 링크 찾기
            blog_links = await page.query_selector_all("a[href*='blog.naver.com'], a[href*='PostView']")
            
            self.log(f"워커 {worker_id}: '{keyword}' - 총 {len(blog_links)}개의 블로그 링크 발견")
            
            seen_urls = set()
            post_containers = []
            
            # 2단계: 각 링크의 상위 컨테이너와 정보 수집
            for link in blog_links:
                try:
                    href = await link.get_attribute("href")
                    if not href or href in seen_urls:
                        continue
                    
                    # 블로그 URL 유효성 검사
                    if not ('blog.naver.com' in href or 'PostView' in href):
                        continue
                    
                    seen_urls.add(href)
                    
                    # 제목 추출
                    try:
                        title = await link.inner_text()
                        title = title.strip() if title else "제목 없음"
                    except:
                        title = "제목 없음"
                    
                    # 컨테이너 위치 정보 수집
                    try:
                        bbox = await link.bounding_box()
                        y_position = bbox['y'] if bbox else 9999
                    except:
                        y_position = 9999
                    
                    # 간단한 광고 필터링
                    try:
                        parent_classes = await link.evaluate("el => el.closest('div')?.className || ''")
                        if any(ad_word in parent_classes.lower() for ad_word in ['ad', 'sponsor', 'shopping', '광고']):
                            continue
                    except:
                        pass
                    
                    post_containers.append({
                        'url': href,
                        'title': title,
                        'y_position': y_position,
                        'link_element': link
                    })
                    
                except Exception as e:
                    continue
            
            # 3단계: Y 좌표로 정렬 (위에서 아래 순서)
            post_containers.sort(key=lambda x: x['y_position'])
            
            # 4단계: 상위 10개 선택
            for container in post_containers[:10]:
                valid_posts.append({
                    'url': container['url'],
                    'title': container['title']
                })
            
            self.log(f"워커 {worker_id}: '{keyword}' - 안정적 방법으로 {len(valid_posts)}개 포스트 추출 완료")
            return valid_posts
            
        except Exception as e:
            self.log(f"워커 {worker_id}: '{keyword}' - 안정적 추출 방법 실패: {str(e)}")
            raise e
    
    async def extract_posts_fallback(self, page, worker_id, keyword):
        """기본 폴백 방식으로 포스트 추출"""
        try:
            valid_posts = []
            
            # 기존 방식의 셀렉터들 시도
            selectors = [
                "a[href*='blog.naver.com']",
                "a[href*='PostView.naver']",
                "a[href*='PostView.nhn']"
            ]
            
            all_links = []
            for selector in selectors:
                try:
                    links = await page.query_selector_all(selector)
                    all_links.extend(links)
                except:
                    continue
            
            seen_urls = set()
            
            for i, link in enumerate(all_links[:15]):  # 최대 15개만 처리
                try:
                    href = await link.get_attribute("href")
                    if not href or href in seen_urls:
                        continue
                    
                    seen_urls.add(href)
                    
                    try:
                        title = await link.inner_text()
                        title = title.strip() if title else f"포스트 {i+1}"
                    except:
                        title = f"포스트 {i+1}"
                    
                    valid_posts.append({
                        'url': href,
                        'title': title
                    })
                    
                    if len(valid_posts) >= 10:
                        break
                        
                except Exception:
                    continue
            
            self.log(f"워커 {worker_id}: '{keyword}' - 폴백 방법으로 {len(valid_posts)}개 포스트 추출")
            return valid_posts
            
        except Exception as e:
            self.log(f"워커 {worker_id}: '{keyword}' - 폴백 방법도 실패: {str(e)}")
            return []
    
    async def crawl_keyword_posts(self, page, keyword, worker_id):
        """키워드별 포스트 상세 정보 크롤링 - 안정성 개선"""
        try:
            crawl_results = []
            
            # 안정적인 방식으로 포스트 추출
            try:
                valid_posts = await self.extract_blog_posts_stable(page, worker_id, keyword)
            except:
                valid_posts = await self.extract_posts_fallback(page, worker_id, keyword)
            
            for idx, post_info in enumerate(valid_posts[:5]):
                try:
                    post_data = {
                        "키워드": keyword,
                        "제목": post_info.get('title', '') if self.extract_title_var.get() else "",
                        "URL": post_info.get('url', ''),
                        "내용": "",
                        "연락처": "",
                        "톡톡링크": ""
                    }
                    
                    # 상세 정보 추출
                    if self.extract_content_var.get() or self.extract_contact_var.get() or self.extract_talktalk_var.get():
                        content, contact, talktalk = await self.extract_post_details(page, post_info.get('url', ''))
                        post_data["내용"] = content if self.extract_content_var.get() else ""
                        post_data["연락처"] = contact if self.extract_contact_var.get() else ""
                        post_data["톡톡링크"] = talktalk if self.extract_talktalk_var.get() else ""
                    
                    crawl_results.append(post_data)
                    
                except Exception as e:
                    self.log(f"워커 {worker_id}: 포스트 크롤링 중 오류: {str(e)}")
            
            return crawl_results
            
        except Exception as e:
            self.log(f"워커 {worker_id}: '{keyword}' 크롤링 중 오류: {str(e)}")
            return []
    
    async def extract_post_details(self, page, url):
        """포스트 상세 정보 추출"""
        content = ""
        contact = ""
        talktalk = ""
        
        try:
            detail_page = await page.context.new_page()
            try:
                await detail_page.goto(url, timeout=30000)
                await detail_page.wait_for_load_state('load', timeout=30000)
                
                # 내용 추출
                if self.extract_content_var.get():
                    content_selectors = [
                        "div.se-main-container",
                        "div[class*='se-main']",
                        "div.post-view",
                        "div.blog-content"
                    ]
                    
                    for c_sel in content_selectors:
                        try:
                            content_elem = await detail_page.query_selector(c_sel)
                            if content_elem:
                                content = await content_elem.inner_text()
                                content = re.sub(r'[^0-9가-힣a-zA-Z.,:~#-?!]', ' ', content)
                                content = re.sub(r' +', ' ', content).strip()[:500]
                                break
                        except:
                            continue
                
                # 연락처 추출 (전화번호, 이메일 등)
                if self.extract_contact_var.get():
                    try:
                        page_text = await detail_page.inner_text("body")
                        # 전화번호 패턴 찾기
                        phone_patterns = [
                            r'(\d{2,3}-\d{3,4}-\d{4})',
                            r'(\d{3}-\d{4}-\d{4})',
                            r'(010-\d{4}-\d{4})',
                            r'(\d{11})'
                        ]
                        for pattern in phone_patterns:
                            matches = re.findall(pattern, page_text)
                            if matches:
                                contact = matches[0]
                                break
                    except:
                        pass
                
                # 톡톡 링크 추출
                if self.extract_talktalk_var.get():
                    try:
                        talktalk_elem = await detail_page.query_selector("a[href*='talk.naver.com']")
                        if talktalk_elem:
                            talktalk = await talktalk_elem.get_attribute("href")
                    except:
                        pass
                
            finally:
                await detail_page.close()
                
        except Exception as e:
            self.log(f"상세 정보 추출 중 오류: {url} - {str(e)}")
        
        return content, contact, talktalk
    
    def save_rank_results(self, rank_results, save_dir):
        """순위 체크 결과 저장 - 저장 옵션에 따라 필터링"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 저장 옵션에 따른 데이터 필터링
        if self.save_option_var.get() == "matched_only":
            # 블로그ID만 저장 - 매칭된 결과만 필터링
            filtered_results = [result for result in rank_results if result['상태'] == 'OK']
            file_prefix = "네이버블로그_매칭결과"
            self.log(f"매칭된 결과만 저장: {len(filtered_results)}개 항목")
        else:
            # 모든 순위 저장 - 전체 결과
            filtered_results = rank_results
            file_prefix = "네이버블로그_전체순위"
            self.log(f"전체 순위 저장: {len(filtered_results)}개 항목")
        
        # 메인 결과 파일 저장
        file_path = os.path.join(save_dir, f"{file_prefix}_{timestamp}.xlsx")
        df = pd.DataFrame(filtered_results)
        df.to_excel(file_path, index=False, engine='openpyxl')
        
        self.log(f"순위 체크 결과가 {file_path}에 저장되었습니다.")
        
        # 요약 통계 생성 (매칭된 결과가 있는 경우에만)
        if filtered_results:
            summary_data = []
            
            # 모든 키워드를 포함하도록 요약 생성 (matched_only와 all_ranks 모두 동일 로직 적용)
            for keyword in set(result['키워드'] for result in rank_results):
                keyword_results = [r for r in rank_results if r['키워드'] == keyword and r['상태'] == 'OK']
                ok_count = len(keyword_results)
                top_rank = min([r['순위'] for r in keyword_results], default=0)
                
                summary_data.append({
                    "키워드": keyword,
                    "매칭_개수": ok_count,
                    "최고_순위": top_rank if top_rank > 0 else "순위권 밖",
                    "1-3위": sum(1 for r in keyword_results if r['순위'] <= 3),
                    "4-6위": sum(1 for r in keyword_results if 4 <= r['순위'] <= 6),
                    "7-10위": sum(1 for r in keyword_results if 7 <= r['순위'] <= 10)
                })
            
            # 요약 파일 저장
            summary_file = os.path.join(save_dir, f"{file_prefix}_요약_{timestamp}.xlsx")
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(summary_file, index=False, engine='openpyxl')
            
            self.log(f"순위 요약 결과가 {summary_file}에 저장되었습니다.")
    
    def save_crawl_results(self, crawl_results, save_dir):
        """크롤링 결과 저장"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        save_format = self.save_format_var.get()
        
        if save_format == "excel":
            file_path = os.path.join(save_dir, f"네이버블로그_크롤링_{timestamp}.xlsx")
            df = pd.DataFrame(crawl_results)
            df.to_excel(file_path, index=False, engine='openpyxl')
        elif save_format == "csv":
            file_path = os.path.join(save_dir, f"네이버블로그_크롤링_{timestamp}.csv")
            df = pd.DataFrame(crawl_results)
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
        elif save_format == "txt":
            file_path = os.path.join(save_dir, f"네이버블로그_크롤링_{timestamp}.txt")
            with open(file_path, 'w', encoding='utf-8') as f:
                for result in crawl_results:
                    for key, value in result.items():
                        f.write(f"{key}: {value}\n")
                    f.write("="*50 + "\n")
        
        self.log(f"크롤링 결과가 {file_path}에 저장되었습니다.")
    
    def update_progress(self, current, total):
        """진행상황 업데이트"""
        progress = int((current / total) * 100)
        self.progress_bar["value"] = progress
        self.progress_var.set(f"진행 중: {progress}% ({current}/{total} 키워드 완료)")
    
    def reset_ui(self):
        """UI 상태 초기화"""
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.progress_var.set("대기 중")

if __name__ == "__main__":
    if hasattr(sys, '_MEIPASS'):
        os.chdir(sys._MEIPASS)
    
    root = tk.Tk()
    app = NaverBlogRankChecker(root)
    root.mainloop()
                        