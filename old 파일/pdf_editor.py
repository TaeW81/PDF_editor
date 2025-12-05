"""
KUNHWA PDF Editor
PDF 문서 편집 및 관리 도구

개발자: 장태웅
버전: 2.0
"""

import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
import fitz
from functools import partial
import os
from PIL import Image, ImageTk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    # tkinterdnd2가 없을 경우를 대비한 fallback
    TkinterDnD = None
    DND_FILES = None

class ModernButton(tk.Button):
    """모던한 디자인의 버튼 클래스"""
    def __init__(self, parent, **kwargs):
        # 툴팁 텍스트 추출
        self.tooltip_text = kwargs.pop('tooltip', None)
        
        # 기본 스타일 설정
        kwargs.setdefault('font', ('Malgun Gothic', 9, 'bold'))
        kwargs.setdefault('bd', 0)
        kwargs.setdefault('relief', 'flat')
        kwargs.setdefault('cursor', 'hand2')
        kwargs.setdefault('activebackground', '#4A90E2')
        kwargs.setdefault('activeforeground', 'white')
        
        super().__init__(parent, **kwargs)
        
        # 마우스 이벤트 바인딩
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)
        self.bind('<Button-1>', self.on_click)
        self.bind('<ButtonRelease-1>', self.on_release)
        
        # 툴팁 관련 변수
        self.tooltip_window = None
        self.tooltip_timer = None
        
        # 초기 상태
        self.is_pressed = False
        self.apply_normal_style()
    
    def apply_normal_style(self):
        """일반 상태 스타일"""
        if self.cget('text') == '빈페이지 삽입':
            self.config(bg='#64B5F6', fg='white')
        elif '회전' in self.cget('text'):
            self.config(bg='#FF9800', fg='white')
        elif 'PDF 열기' in self.cget('text'):
            self.config(bg='#1976D2', fg='white')
        elif 'PDF 병합' in self.cget('text'):
            self.config(bg='#42A5F5', fg='white')
        elif 'PDF 저장' in self.cget('text'):
            self.config(bg='#2196F3', fg='white')
        elif '페이지에 맞춤' in self.cget('text'):
            self.config(bg='#4CAF50', fg='white')
        else:
            self.config(bg='#607D8B', fg='white')
    
    def on_enter(self, event):
        """마우스 진입 시 호버 효과 및 툴팁 표시"""
        if not self.is_pressed:
            if self.cget('text') == '빈페이지 삽입':
                self.config(bg='#42A5F5')
            elif '회전' in self.cget('text'):
                self.config(bg='#FFB74D')
            elif 'PDF 열기' in self.cget('text'):
                self.config(bg='#1565C0')
            elif 'PDF 병합' in self.cget('text'):
                self.config(bg='#1E88E5')
            elif 'PDF 저장' in self.cget('text'):
                self.config(bg='#1976D2')
            elif '페이지에 맞춤' in self.cget('text'):
                self.config(bg='#66BB6A')
            else:
                self.config(bg='#78909C')
        
        # 툴팁 표시 (0.5초 후)
        if self.tooltip_text:
            self.tooltip_timer = self.after(500, self.show_tooltip)
    
    def on_leave(self, event):
        """마우스 이탈 시 원래 스타일 및 툴팁 숨김"""
        if not self.is_pressed:
            self.apply_normal_style()
        
        # 툴팁 타이머 취소 및 툴팁 숨김
        if self.tooltip_timer:
            self.after_cancel(self.tooltip_timer)
            self.tooltip_timer = None
        
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
    
    def show_tooltip(self):
        """툴팁 창 표시"""
        if not self.tooltip_text:
            return
        
        # 툴팁 창 생성
        self.tooltip_window = tk.Toplevel(self)
        self.tooltip_window.overrideredirect(True)
        self.tooltip_window.configure(bg='#333333')
        
        # 툴팁 내용
        tooltip_label = tk.Label(
            self.tooltip_window,
            text=self.tooltip_text,
            bg='#333333',
            fg='white',
            font=('Malgun Gothic', 9),
            wraplength=300,
            justify=tk.LEFT,
            padx=10,
            pady=8
        )
        tooltip_label.pack()
        
        # 툴팁 위치 계산
        x = self.winfo_rootx() + self.winfo_width() // 2
        y = self.winfo_rooty() - 10
        
        if y < 0:
            y = self.winfo_rooty() + self.winfo_height() + 10
        
        self.tooltip_window.geometry(f"+{x}+{y}")
        self.tooltip_window.lift()
    
    def on_click(self, event):
        """클릭 시 눌림 효과"""
        self.is_pressed = True
        self.config(relief='sunken')
    
    def on_release(self, event):
        """클릭 해제 시 원래 스타일"""
        self.is_pressed = False
        self.config(relief='flat')
        self.apply_normal_style()

class PDFEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KUNHWA PDF Editor")
        
        # 변수 초기화
        self.doc = None
        self.thumbnail_frames = []
        self.selected_indices = set()
        self.dragged_frame = None
        self.drag_start_y = 0
        self.thumb_scale = 1.0
        self.preview_scale = 1.0
        
        # 색상 및 폰트 설정
        self.colors = {
            'primary': '#1976D2',
            'secondary': '#42A5F5',
            'accent': '#FF9800',
            'success': '#4CAF50',
            'warning': '#FF9800',
            'error': '#F44336',
            'background': '#f0f0f0',
            'surface': '#ffffff',
            'text': '#333333',
            'text_secondary': '#666666'
        }
        
        self.fonts = {
            'title': ('Malgun Gothic', 16, 'bold'),
            'heading': ('Malgun Gothic', 12, 'bold'),
            'body': ('Malgun Gothic', 10),
            'button': ('Malgun Gothic', 9, 'bold'),
            'caption': ('Malgun Gothic', 8)
        }
        
        # UI 설정
        self.setup_ui()
        
        # tkinterdnd2 지원 확인 및 설정
        if TkinterDnD and hasattr(root, 'drop_target_register'):
            self.dnd_supported = True
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_drop_files)
            self.root.dnd_bind('<<DragEnter>>', self.on_drag_enter)
            self.root.dnd_bind('<<DragLeave>>', self.on_drag_leave)
        else:
            self.dnd_supported = False
        
        # Delete 키 바인딩
        self.root.bind_all("<Delete>", lambda e: self.delete_pages())
        
        # 로고 로딩
        self.load_logo()
        
        # 초기 로고 표시
        self.root.after(100, self.show_logo)
    
    def setup_ui(self):
        """UI 초기 설정"""
        # 상단 버튼 프레임
        top_frame = tk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        # 왼쪽 버튼 그룹 (PDF 관련)
        left_buttons = tk.Frame(top_frame)
        left_buttons.pack(side=tk.LEFT, padx=(10, 0))
        
        ModernButton(left_buttons, text="PDF 열기", command=self.open_pdf, width=10, height=2, tooltip="PDF 문서를 엽니다.").pack(side=tk.LEFT, padx=(0, 5))
        ModernButton(left_buttons, text="PDF 병합", command=self.merge_pdf_files, width=10, height=2, tooltip="PDF 문서를 병합합니다.").pack(side=tk.LEFT, padx=(0, 5))
        ModernButton(left_buttons, text="PDF 저장", command=self.save_pdf, width=10, height=2, tooltip="PDF 문서를 저장합니다.").pack(side=tk.LEFT, padx=(0, 5))
        
        # 첫 번째 구분선
        separator1 = tk.Frame(top_frame, width=1, bg='#CCCCCC')
        separator1.pack(side=tk.LEFT, fill=tk.Y, padx=15, pady=5)
        
        # 중앙 버튼 그룹 (회전 + 삽입 + 우수흐름)
        center_buttons = tk.Frame(top_frame)
        center_buttons.pack(side=tk.LEFT, padx=0)
        
        ModernButton(center_buttons, text="우로 회전", command=lambda: self.rotate_pages(90), width=10, height=2, tooltip="페이지를 우측으로 회전합니다.").pack(side=tk.LEFT, padx=(0, 5))
        ModernButton(center_buttons, text="좌로 회전", command=lambda: self.rotate_pages(-90), width=10, height=2, tooltip="페이지를 좌측으로 회전합니다.").pack(side=tk.LEFT, padx=(0, 5))
        
        # 회전과 삽입 사이 구분선
        separator_rotate_insert = tk.Frame(center_buttons, width=1, bg='#CCCCCC')
        separator_rotate_insert.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        
        ModernButton(center_buttons, text="빈페이지 삽입", command=self.show_insert_menu, width=12, height=2, tooltip="새로운 빈 페이지를 삽입합니다.").pack(side=tk.LEFT, padx=(5, 0))
        

        
        # 두 번째 구분선
        separator2 = tk.Frame(top_frame, width=1, bg='#CCCCCC')
        separator2.pack(side=tk.LEFT, fill=tk.Y, padx=15, pady=5)
        
        # 오른쪽 버튼 그룹
        right_buttons = tk.Frame(top_frame)
        right_buttons.pack(side=tk.RIGHT, padx=(0, 10))
        
        ModernButton(right_buttons, text="페이지에 맞춤", command=self.fit_to_page, width=14, height=2, tooltip="페이지를 창 크기에 맞춥니다.").pack(side=tk.LEFT, padx=0)
        
        # 좌우 분할 컨테이너
        self.content_frame = tk.Frame(self.root)
        self.content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # 왼쪽: 썸네일 영역
        left_container = tk.Frame(self.content_frame)
        left_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(left_container, bg="white")
        self.scrollbar = tk.Scrollbar(left_container, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)
        self.default_frame_bg = self.scrollable_frame.cget("bg")
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 오른쪽: 선택 페이지 큰 미리보기 영역
        self.preview_frame = tk.Frame(self.content_frame, bg="#E5E5E5")
        self.preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 미리보기 캔버스와 스크롤바
        preview_container = tk.Frame(self.preview_frame)
        preview_container.pack(fill=tk.BOTH, expand=True)
        
        self.preview_canvas = tk.Canvas(preview_container, bg="#E5E5E5", highlightthickness=0)
        self.preview_h_scrollbar = tk.Scrollbar(preview_container, orient=tk.HORIZONTAL, command=self.preview_canvas.xview)
        self.preview_v_scrollbar = tk.Scrollbar(preview_container, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        
        self.preview_canvas.configure(xscrollcommand=self.preview_h_scrollbar.set, yscrollcommand=self.preview_v_scrollbar.set)
        
        # 스크롤바 배치
        self.preview_h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 마우스 휠 이벤트 바인딩 (독립적 확대/축소)
        self.canvas.bind('<MouseWheel>', self.on_thumbnail_scroll)
        self.canvas.bind('<Control-MouseWheel>', self.on_thumbnail_scroll)
        # Windows에서 Ctrl+마우스 휠 이벤트를 더 명확하게 바인딩
        self.canvas.bind('<Control-Button-4>', lambda e: self.on_thumbnail_zoom(e, True))   # 휠 위로
        self.canvas.bind('<Control-Button-5>', lambda e: self.on_thumbnail_zoom(e, False))  # 휠 아래로
        
        self.preview_canvas.bind('<MouseWheel>', self.on_preview_scroll)
        self.preview_canvas.bind('<Control-MouseWheel>', self.on_preview_scroll)
        
        # 창 크기 변경 이벤트 바인딩
        self.root.bind('<Configure>', self.on_window_resize)
        
        # 선택 상태 색상 설정
        self.selection_color = '#FF0000'  # 빨간색
        self.selection_bg = '#FFE6E6'     # 연한 빨간색 배경
        self.selection_highlight = '#FF0000'  # 빨간색 테두리
        
        # 저작권 표시
        copyright_frame = tk.Frame(self.root, bg='#f0f0f0')
        copyright_frame.pack(side=tk.BOTTOM, fill=tk.X)
        copyright_label = tk.Label(
            copyright_frame,
            text="© 2025 Kunhwa Engineering & Consulting | Developed by TaeWoong Jang",
            font=('Malgun Gothic', 8),
            bg='#f0f0f0',
            fg='#666666'
        )
        copyright_label.pack(expand=True, pady=5)
    
    def load_logo(self):
        """KUNHWA 로고 로딩"""
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            logo_path = os.path.join(script_dir, "data", "kunhwa_logo.png")
            
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((200, 100), Image.Resampling.LANCZOS)
                self.logo_photo = logo_img
                self.logo_photo_tk = ImageTk.PhotoImage(logo_img)
            else:
                self.logo_photo = None
                self.logo_photo_tk = None
        except Exception as e:
            self.logo_photo = None
            self.logo_photo_tk = None
    
    def show_logo(self):
        """로고 표시 (PDF가 열려있지 않을 때만)"""
        if not self.doc and self.logo_photo_tk:
            try:
                self.preview_canvas.delete("all")
                canvas_width = self.preview_canvas.winfo_width()
                canvas_height = self.preview_canvas.winfo_height()
                
                if canvas_width > 1 and canvas_height > 1:
                    # 로고를 캔버스 중앙에 배치
                    logo_x = (canvas_width - 200) // 2
                    logo_y = (canvas_height - 100) // 2
                    self.preview_canvas.create_image(logo_x, logo_y, anchor=tk.NW, image=self.logo_photo_tk)
            except Exception as e:
                pass
    
    def open_pdf(self):
        """PDF 파일 열기"""
        file_path = filedialog.askopenfilename(
            title="PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            try:
                self.doc = fitz.open(file_path)
                self.refresh_thumbnails()
                self.update_preview()
            except Exception as e:
                error_msg = f"PDF 파일을 열 수 없습니다.\n\n오류 상세:\n{str(e)}\n\n파일 경로: {file_path}"
                messagebox.showerror("오류", error_msg)
                print(f"PDF 열기 오류: {e}")
                print(f"파일 경로: {file_path}")
                import traceback
                traceback.print_exc()
    
    def refresh_thumbnails_with_scale(self):
        """크기가 조정된 썸네일 새로고침"""
        if not self.doc:
            return
        
        try:
            print(f"썸네일 새로고침 시작 (스케일: {self.thumb_scale:.2f})")
            
            # 기존 썸네일 제거
            for frame in self.thumbnail_frames:
                try:
                    frame.destroy()
                except Exception as e:
                    continue
            
            self.thumbnail_frames.clear()
            self.selected_indices.clear()
            
            # 새 썸네일 생성 (현재 스케일 적용)
            for i in range(len(self.doc)):
                try:
                    self.create_thumbnail(i)
                except Exception as e:
                    continue
            
            print(f"썸네일 {len(self.thumbnail_frames)}개 생성 완료")
            
            # 그리드 레이아웃 업데이트
            self.root.after(100, self.update_grid_layout)
            
            # 첫 번째 페이지 선택
            if self.thumbnail_frames:
                self.handle_selection(0, True, None)
                
        except Exception as e:
            print(f"크기 조정 썸네일 새로고침 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def refresh_thumbnails(self):
        """썸네일 새로고침"""
        if not self.doc:
            return
        
        try:
            # 기존 썸네일 제거
            for frame in self.thumbnail_frames:
                try:
                    frame.destroy()
                except Exception as e:
                    continue
            
            self.thumbnail_frames.clear()
            self.selected_indices.clear()
            
            # 새 썸네일 생성
            for i in range(len(self.doc)):
                try:
                    self.create_thumbnail(i)
                except Exception as e:
                    continue
            
            # 그리드 레이아웃 업데이트 (더 긴 지연 후)
            self.root.after(500, self.update_grid_layout)
            
            # 첫 번째 페이지 선택
            if self.thumbnail_frames:
                self.handle_selection(0, True, None)
                
        except Exception as e:
            print(f"썸네일 새로고침 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def create_thumbnail(self, page_index):
        """썸네일 생성"""
        try:
            page = self.doc[page_index]
            pix = page.get_pixmap(matrix=fitz.Matrix(0.3 * self.thumb_scale, 0.3 * self.thumb_scale))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # 페이지 방향에 따른 썸네일 크기 조정
            img_ratio = img.width / img.height
            
            if img_ratio > 1.0:  # 가로형 페이지
                # 가로형: 가로를 기준으로 크기 조정
                target_width = int(160 * self.thumb_scale)
                target_height = int(target_width / img_ratio)
                # 최소 높이 보장
                if target_height < int(120 * self.thumb_scale):
                    target_height = int(120 * self.thumb_scale)
                    target_width = int(target_height * img_ratio)
            else:  # 세로형 페이지
                # 세로형: 세로를 기준으로 크기 조정
                target_height = int(160 * self.thumb_scale)
                target_width = int(target_height * img_ratio)
                # 최소 너비 보장
                if target_width < int(120 * self.thumb_scale):
                    target_width = int(120 * self.thumb_scale)
                    target_height = int(target_width / img_ratio)
            
            # 이미지 리사이즈
            img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
            
            # PDF 내용만 표시 (배경 없음)
            photo = ImageTk.PhotoImage(img)
            
            # 썸네일 프레임 생성 (테두리 없음)
            frame = tk.Frame(
                self.scrollable_frame, 
                bg='white', 
                relief=tk.FLAT,  # 테두리 없음
                bd=0,            # 테두리 두께 0
                highlightthickness=0,
                cursor='hand2'  # 마우스 커서를 손가락 모양으로 변경
            )
            
            # PDF 이미지만 표시 (테두리 포함)
            image_label = tk.Label(
                frame, 
                image=photo, 
                bg='white',
                relief=tk.SOLID,  # 이미지에만 테두리
                bd=1              # 이미지 테두리 두께
            )
            image_label.pack(padx=2, pady=(2, 0))
            image_label.image = photo
            
            # 페이지 번호 별도 라벨 (테두리 없음)
            page_label = tk.Label(
                frame,
                text=f"페이지 {page_index + 1}",
                bg='white',
                fg='black',
                font=('Malgun Gothic', 9),
                relief=tk.FLAT,  # 테두리 없음
                bd=0             # 테두리 두께 0
            )
            page_label.pack(pady=(0, 2))
            
            # 이벤트 바인딩 (이미지 라벨에만)
            image_label.bind('<Button-1>', lambda e, idx=page_index: self.handle_selection(idx, False, e))
            image_label.bind('<Double-Button-1>', lambda e, idx=page_index: self.on_double_click(idx))
            image_label.bind('<B1-Motion>', lambda e, idx=page_index: self.on_drag(e, idx))
            image_label.bind('<ButtonRelease-1>', lambda e, idx=page_index: self.on_drop(e, idx))
            image_label.bind('<Enter>', lambda e, f=frame: self.on_enter(f))
            image_label.bind('<Leave>', lambda e, f=frame: self.on_leave(f))
            
            # 프레임을 먼저 pack으로 배치 (나중에 grid로 변경)
            frame.pack(fill=tk.X, padx=2, pady=2)
            
            self.thumbnail_frames.append(frame)
            
        except Exception as e:
            print(f"썸네일 생성 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def handle_selection(self, index, clear_others=False, event=None):
        """선택 처리"""
        # 이벤트가 있는 경우 키 상태 확인
        if event:
            ctrl_pressed = event.state & 0x4  # Ctrl 키
            shift_pressed = event.state & 0x1  # Shift 키
            
            if ctrl_pressed:
                # Ctrl 키: 토글 선택 (기존 선택 유지)
                if index in self.selected_indices:
                    self.selected_indices.remove(index)
                else:
                    self.selected_indices.add(index)
            elif shift_pressed and self.selected_indices:
                # Shift 키: 범위 선택
                last_selected = max(self.selected_indices)
                start_idx = min(last_selected, index)
                end_idx = max(last_selected, index)
                
                # 기존 선택 초기화
                self.selected_indices.clear()
                
                # 범위 내 모든 페이지 선택
                for i in range(start_idx, end_idx + 1):
                    self.selected_indices.add(i)
            else:
                # 일반 클릭: 단일 선택 (기존 선택 모두 초기화)
                self.selected_indices.clear()
                self.selected_indices.add(index)
        else:
            # 이벤트가 없는 경우 (프로그램 내부 호출)
            if clear_others:
                self.selected_indices.clear()
            
            # 단일 선택으로 변경
            self.selected_indices.clear()
            self.selected_indices.add(index)
        
        self.update_selection_highlight()
        self.update_preview()
        
        # 선택 상태는 내부적으로만 관리 (콘솔 출력 없음)
    
    def update_selection_highlight(self):
        """선택 상태 하이라이트 업데이트 - 시각적 표시 없음"""
        try:
            for i, frame in enumerate(self.thumbnail_frames):
                try:
                    # 모든 썸네일을 동일한 스타일로 설정 (선택 상태와 무관)
                    frame.config(
                        bg='white', 
                        highlightbackground='black',  # 빈 문자열 대신 검은색 사용
                        highlightthickness=0,
                        relief=tk.FLAT,  # 테두리 없음
                        bd=0             # 테두리 두께 0
                    )
                    
                    # 프레임 내의 라벨도 동일한 배경색으로 설정
                    for child in frame.winfo_children():
                        if isinstance(child, tk.Label):
                            child.config(bg='white')
                except Exception as e:
                    print(f"썸네일 {i+1} 스타일 설정 오류: {e}")
                    continue
        except Exception as e:
            print(f"썸네일 스타일 설정 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def on_enter(self, frame):
        """마우스 진입"""
        # 마우스 오버 시 시각적 변화 없음 (선택된 상태만 유지)
        pass
    
    def on_leave(self, frame):
        """마우스 이탈"""
        # 마우스 이탈 시 시각적 변화 없음 (선택된 상태만 유지)
        pass
    
    def update_preview(self):
        """미리보기 업데이트"""
        if not self.doc or not self.selected_indices:
            self.show_logo()
            return
        
        try:
            # 첫 번째 선택된 페이지 표시
            page_index = min(self.selected_indices)
            page = self.doc[page_index]
            
            # 페이지 렌더링
            matrix = fitz.Matrix(self.preview_scale, self.preview_scale)
            pix = page.get_pixmap(matrix=matrix)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # 캔버스 크기에 맞게 조정
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:
                # 이미지가 캔버스보다 클 경우 스크롤 가능하도록 설정
                self.preview_canvas.delete("all")
                self.preview_image_tk = ImageTk.PhotoImage(img)
                
                # 캔버스 스크롤 영역 설정
                self.preview_canvas.configure(scrollregion=(0, 0, pix.width, pix.height))
                
                # 이미지 표시
                self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=self.preview_image_tk)
                
        except Exception as e:
            print(f"미리보기 업데이트 오류: {e}")
    
    def update_grid_layout(self):
        """썸네일 그리드 레이아웃 업데이트"""
        if not self.doc or not self.thumbnail_frames:
            return
        
        try:
            self.canvas.update_idletasks()
            canvas_width = self.canvas.winfo_width()
            
            if canvas_width < 20:
                # 캔버스가 아직 준비되지 않았으면 잠시 후 다시 시도
                self.root.after(50, self.update_grid_layout)
                return
            
            # 크기가 조정된 썸네일에 맞춘 그리드 계산
            # 기본 썸네일 너비 (가로형 기준, 여백 포함)
            base_thumb_width = int(170 * self.thumb_scale)  # 160 * 스케일 + 10 (여백)
            base_thumb_height = int(180 * self.thumb_scale)  # 160 * 스케일 + 20 (텍스트 및 여백)
            
            # 한 줄에 들어갈 수 있는 썸네일 개수 계산
            col_count = max(canvas_width // base_thumb_width, 1)
            
            # 모든 프레임을 먼저 pack에서 제거
            for frame in self.thumbnail_frames:
                try:
                    frame.pack_forget()
                except Exception as e:
                    continue
            
            # 그리드 배치
            for i, frame in enumerate(self.thumbnail_frames):
                try:
                    row = i // col_count
                    col = i % col_count
                    
                    # grid로 배치
                    frame.grid(row=row, column=col, padx=3, pady=3, sticky="nsew")
                except Exception as e:
                    continue
            
            # 컬럼 가중치 설정
            for col in range(col_count):
                try:
                    self.scrollable_frame.grid_columnconfigure(col, weight=1)
                except Exception as e:
                    continue
            
            # 스크롤 영역 업데이트
            self.scrollable_frame.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            
        except Exception as e:
            print(f"그리드 레이아웃 업데이트 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def on_thumbnail_zoom(self, event, zoom_in):
        """Windows에서 Ctrl+마우스 휠 줌 처리"""
        try:
            old_scale = self.thumb_scale
            
            if zoom_in:
                self.thumb_scale = min(self.thumb_scale * 1.1, 3.0)  # 최대 3배
            else:
                self.thumb_scale = max(self.thumb_scale * 0.9, 0.3)  # 최소 0.3배
            
            # 스케일이 변경된 경우에만 새로고침
            if abs(self.thumb_scale - old_scale) > 0.01:
                print(f"썸네일 크기 조정 (Windows): {old_scale:.2f} -> {self.thumb_scale:.2f}")
                # 썸네일 크기 변경 후 그리드 레이아웃 재조정
                self.refresh_thumbnails_with_scale()
                
        except Exception as e:
            print(f"썸네일 줌 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def on_thumbnail_scroll(self, event):
        """썸네일 영역 스크롤/줌"""
        try:
            # Ctrl 키 상태 확인 (Windows)
            ctrl_pressed = event.state & 0x4  # Ctrl 키
            
            if ctrl_pressed:
                # 줌 기능: 썸네일 크기 조정
                old_scale = self.thumb_scale
                
                if event.delta > 0:
                    self.thumb_scale = min(self.thumb_scale * 1.1, 3.0)  # 최대 3배
                else:
                    self.thumb_scale = max(self.thumb_scale * 0.9, 0.3)  # 최소 0.3배
                
                # 스케일이 변경된 경우에만 새로고침
                if abs(self.thumb_scale - old_scale) > 0.01:
                    print(f"썸네일 크기 조정: {old_scale:.2f} -> {self.thumb_scale:.2f}")
                    # 썸네일 크기 변경 후 그리드 레이아웃 재조정
                    self.refresh_thumbnails_with_scale()
            else:
                # 스크롤
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                
        except Exception as e:
            print(f"썸네일 스크롤/줌 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def on_preview_scroll(self, event):
        """미리보기 영역 스크롤/줌"""
        if event.state & 0x4:  # Ctrl 키가 눌린 상태
            # 줌
            if event.delta > 0:
                self.preview_scale = min(self.preview_scale * 1.1, 5.0)
            else:
                self.preview_scale = max(self.preview_scale * 0.9, 0.1)
            self.update_preview()
        else:
            # 스크롤
            self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def fit_to_page(self):
        """페이지에 맞춤"""
        self.preview_scale = 1.0
        self.update_preview()
    
    def rotate_pages(self, degrees):
        """페이지 회전"""
        if not self.doc or not self.selected_indices:
            messagebox.showwarning("경고", "회전할 페이지를 선택해주세요.")
            return
        
        try:
            for index in self.selected_indices:
                page = self.doc[index]
                page.set_rotation(page.rotation + degrees)
            
            self.refresh_thumbnails()
            messagebox.showinfo("완료", f"선택된 페이지를 {degrees}도 회전했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"페이지 회전 중 오류가 발생했습니다.\n{str(e)}")
    
    def delete_pages(self):
        """선택된 페이지 삭제"""
        if not self.doc or not self.selected_indices:
            return
        
        if messagebox.askyesno("확인", "선택된 페이지를 삭제하시겠습니까?"):
            try:
                # 인덱스를 내림차순으로 정렬하여 뒤에서부터 삭제
                indices_to_delete = sorted(self.selected_indices, reverse=True)
                for index in indices_to_delete:
                    self.doc.delete_page(index)
                
                self.selected_indices.clear()
                self.refresh_thumbnails()
                messagebox.showinfo("완료", "선택된 페이지가 삭제되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"페이지 삭제 중 오류가 발생했습니다.\n{str(e)}")
    
    def show_insert_menu(self):
        """빈페이지 삽입 메뉴"""
        popup = tk.Toplevel(self.root)
        popup.title("빈페이지 삽입")
        popup.geometry("200x200")
        popup.resizable(False, False)
        
        # 팝업을 메인 창 중앙에 위치
        popup.transient(self.root)
        popup.grab_set()
        
        # 메인 프레임
        main_frame = tk.Frame(popup, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # 제목
        title_label = tk.Label(
            main_frame,
            text="빈페이지 삽입",
            font=('Malgun Gothic', 12, 'bold'),
            bg='#f0f0f0',
            fg='#333333'
        )
        title_label.pack(pady=(0, 15))
        
        # 버튼들
        ModernButton(main_frame, text="A4 가로", command=lambda: self.insert_blank_page("a4_h", popup), width=12, height=2).pack(fill=tk.X, padx=15, pady=6)
        ModernButton(main_frame, text="A4 세로", command=lambda: self.insert_blank_page("a4_v", popup), width=12, height=2).pack(fill=tk.X, padx=15, pady=6)
        ModernButton(main_frame, text="A3 가로", command=lambda: self.insert_blank_page("a3_h", popup), width=12, height=2).pack(fill=tk.X, padx=15, pady=6)
        ModernButton(main_frame, text="A3 세로", command=lambda: self.insert_blank_page("a3_v", popup), width=12, height=2).pack(fill=tk.X, padx=15, pady=6)
        
        # 닫기 버튼
        ModernButton(main_frame, text="닫기", command=popup.destroy, width=12, height=2).pack(fill=tk.X, padx=15, pady=(10, 15))
        
        # 팝업을 화면 중앙에 위치
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() // 2) - (popup.winfo_width() // 2)
        y = (popup.winfo_screenheight() // 2) - (popup.winfo_height() // 2)
        popup.geometry(f"+{x}+{y}")
        
        # 포커스 설정
        popup.focus_set()
    
    def insert_blank_page(self, page_type, popup):
        """빈페이지 삽입"""
        try:
            # 페이지 크기 설정
            if page_type == "a4_h":
                width, height = 595, 420  # A4 가로
            elif page_type == "a4_v":
                width, height = 420, 595  # A4 세로
            elif page_type == "a3_h":
                width, height = 842, 595  # A3 가로
            elif page_type == "a3_v":
                width, height = 595, 842  # A3 세로
            else:
                return
            
            # 새 페이지 생성
            if self.doc:
                page = self.doc.new_page(width=width, height=height)
            else:
                self.doc = fitz.open()
                page = self.doc.new_page(width=width, height=height)
            
            popup.destroy()
            self.refresh_thumbnails()
            messagebox.showinfo("완료", "빈페이지가 삽입되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"빈페이지 삽입 중 오류가 발생했습니다.\n{str(e)}")
    

    
    def on_double_click(self, index):
        """더블클릭 이벤트"""
        if index in self.selected_indices:
            self.move_pages_to_position(index)
    
    def move_pages_to_position(self, target_index):
        """선택된 페이지를 특정 위치로 이동"""
        if not self.selected_indices or len(self.selected_indices) <= 1:
            return
        
        # 이동할 페이지들
        pages_to_move = list(self.selected_indices)
        pages_to_move.remove(target_index)
        
        try:
            # 페이지 이동
            for page_index in pages_to_move:
                if page_index < target_index:
                    # 앞에서 뒤로 이동
                    self.doc.move_page(page_index, target_index - 1)
                else:
                    # 뒤에서 앞으로 이동
                    self.doc.move_page(page_index, target_index)
            
            self.selected_indices.clear()
            self.selected_indices.add(target_index)
            self.refresh_thumbnails()
            messagebox.showinfo("완료", "페이지가 이동되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"페이지 이동 중 오류가 발생했습니다.\n{str(e)}")
    
    def on_drag(self, event, index):
        """드래그 시작"""
        if index in self.selected_indices:
            self.dragged_frame = self.thumbnail_frames[index]
            self.drag_start_y = event.y_root
    
    def on_drop(self, event, index):
        """드롭 처리"""
        if self.dragged_frame and self.dragged_frame != self.thumbnail_frames[index]:
            try:
                # 드래그된 페이지의 원래 인덱스 찾기
                dragged_index = self.thumbnail_frames.index(self.dragged_frame)
                
                # 페이지 이동
                if dragged_index < index:
                    self.doc.move_page(dragged_index, index - 1)
                else:
                    self.doc.move_page(dragged_index, index)
                
                self.dragged_frame = None
                self.refresh_thumbnails()
                
            except Exception as e:
                messagebox.showerror("오류", f"페이지 이동 중 오류가 발생했습니다.\n{str(e)}")
    
    def merge_pdf_files(self):
        """PDF 파일 병합"""
        file_paths = filedialog.askopenfilenames(
            title="병합할 PDF 파일들 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if not file_paths:
            return
        
        try:
            # 새 문서 생성 또는 기존 문서에 병합
            if not self.doc:
                self.doc = fitz.open(file_paths[0])
                file_paths = file_paths[1:]
            
            # 나머지 파일들 병합
            for file_path in file_paths:
                other_doc = fitz.open(file_path)
                self.doc.insert_pdf(other_doc)
                other_doc.close()
            
            self.refresh_thumbnails()
            messagebox.showinfo("완료", "PDF 파일이 병합되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"PDF 병합 중 오류가 발생했습니다.\n{str(e)}")
    
    def save_pdf(self):
        """PDF 저장"""
        if not self.doc:
            messagebox.showwarning("경고", "저장할 PDF가 없습니다.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="PDF 저장",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.doc.save(file_path)
                messagebox.showinfo("완료", "PDF가 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"PDF 저장 중 오류가 발생했습니다.\n{str(e)}")
    
    def on_drop_files(self, event):
        """외부 파일 드롭 처리"""
        if not self.dnd_supported:
            return
        
        try:
            # 드롭된 파일 경로 파싱
            data = event.data
            # 다양한 형식의 경로 처리
            if data.startswith('{') and data.endswith('}'):
                # 중괄호로 감싸진 경로
                data = data[1:-1]
            
            # 따옴표 제거
            data = data.strip('"')
            
            # 파일 경로 추출
            import re
            file_paths = re.findall(r'[^;]+', data)
            
            # PDF 파일만 필터링
            pdf_files = []
            for path in file_paths:
                path = path.strip()
                if path.lower().endswith('.pdf'):
                    # 절대 경로로 변환
                    abs_path = os.path.abspath(path)
                    if os.path.exists(abs_path):
                        pdf_files.append(abs_path)
                    else:
                        print(f"파일이 존재하지 않음: {abs_path}")
                else:
                    print(f"PDF 파일이 아님: {path}")
            
            if not pdf_files:
                messagebox.showwarning("경고", "유효한 PDF 파일을 찾을 수 없습니다.")
                return
            
            # 삽입 위치: 마우스 드롭 위치 기반으로 계산
            drop_index = self.get_drop_index_from_position(event.x_root, event.y_root)
            
            if drop_index is None:
                # 위치를 찾을 수 없으면 문서 끝에 추가
                drop_index = len(self.doc)
            
            # 병합 실행
            success_count = 0
            running_index = drop_index
            
            for pdf_path in pdf_files:
                try:
                    added = self.merge_external_pdf(pdf_path, running_index)
                    if added > 0:
                        running_index += added
                        success_count += 1
                except Exception as e:
                    print(f"병합 중 오류 발생: {pdf_path}, 오류: {e}")
                    messagebox.showwarning("병합 실패", f"병합에 실패했습니다.\n{pdf_path}\n{e}")
            
            if success_count > 0:
                # 병합 완료 시 썸네일 새로고침
                self.refresh_thumbnails()
                self.update_preview()
                # 성공 메시지 상자 제거
                # messagebox.showinfo("병합 완료", f"{success_count}개 파일이 성공적으로 병합되었습니다.")
            else:
                messagebox.showwarning("병합 실패", "모든 파일 병합에 실패했습니다.")
                
        except Exception as e:
            print(f"드롭 파일 처리 중 오류: {e}")
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다.\n{str(e)}")
    
    def get_drop_index_from_position(self, x, y):
        """마우스 드롭 위치를 기반으로 삽입할 페이지 인덱스를 계산"""
        if not self.doc or not self.thumbnail_frames:
            return 0
        
        try:
            # 왼쪽 썸네일 영역의 위치 계산
            left_frame_x = self.content_frame.winfo_rootx()
            left_frame_y = self.content_frame.winfo_rooty()
            left_frame_w = self.canvas.winfo_width()
            left_frame_h = self.canvas.winfo_height()
            
            # 마우스가 왼쪽 썸네일 영역에 있는지 확인
            if (x < left_frame_x or x > left_frame_x + left_frame_w or 
                y < left_frame_y or y > left_frame_y + left_frame_h):
                return None
            
            # 썸네일 프레임들의 위치와 비교하여 삽입 위치 결정
            for i, frame in enumerate(self.thumbnail_frames):
                try:
                    frame_x = frame.winfo_rootx()
                    frame_y = frame.winfo_rooty()
                    frame_w = frame.winfo_width()
                    frame_h = frame.winfo_height()
                    
                    # 마우스가 이 썸네일 위에 있는지 확인
                    if (frame_x <= x <= frame_x + frame_w and 
                        frame_y <= y <= frame_y + frame_h):
                        return i
                    
                    # 마우스가 썸네일들 사이에 있는 경우 (세로 방향)
                    if i > 0:
                        prev_frame = self.thumbnail_frames[i-1]
                        prev_frame_bottom = prev_frame.winfo_rooty() + prev_frame.winfo_height()
                        
                        if (prev_frame_bottom <= y <= frame_y and 
                            frame_x <= x <= frame_x + frame_w):
                            return i
                    
                except Exception as e:
                    continue
            
            # 마우스가 마지막 썸네일 아래쪽에 있으면 문서 끝에 추가
            if self.thumbnail_frames:
                last_frame = self.thumbnail_frames[-1]
                try:
                    last_frame_bottom = last_frame.winfo_rooty() + last_frame.winfo_height()
                    if y > last_frame_bottom:
                        return len(self.doc)
                except Exception as e:
                    pass
            
            # 마우스가 첫 번째 썸네일 위쪽에 있으면 첫 번째 위치에 추가
            if self.thumbnail_frames:
                first_frame = self.thumbnail_frames[0]
                try:
                    first_frame_top = first_frame.winfo_rooty()
                    if y < first_frame_top:
                        return 0
                except Exception as e:
                    pass
            
            # 기본값은 첫 번째 위치
            return 0
            
        except Exception as e:
            print(f"드롭 위치 계산 중 오류: {e}")
            return 0
    
    def merge_external_pdf(self, pdf_path, insert_index):
        """외부 PDF 파일을 특정 위치에 병합"""
        try:
            if not self.doc:
                self.doc = fitz.open()
            
            external_doc = fitz.open(pdf_path)
            page_count = len(external_doc)
            
            if page_count == 0:
                external_doc.close()
                return 0
            
            # PyMuPDF 버전에 따른 insert_pdf 처리
            try:
                # PyMuPDF 1.26.3dldi 호환성을 위한 처리
                if insert_index >= len(self.doc):
                    # 문서 끝에 추가
                    self.doc.insert_pdf(external_doc)
                else:
                    # 특정 위치에 삽입 시도
                    try:
                        self.doc.insert_pdf(external_doc, start_at=insert_index)
                    except:
                        # insert_pdf가 실패하면 끝에 추가 후 이동
                        self.doc.insert_pdf(external_doc)
                        # 추가된 페이지들을 원하는 위치로 이동
                        for i in range(page_count):
                            self.doc.move_page(len(self.doc) - 1, insert_index + i)
            except Exception as e:
                print(f"insert_pdf 실패, 대체 방법 사용: {e}")
                # 대체 방법: 끝에 추가 후 이동
                self.doc.insert_pdf(external_doc)
                for i in range(page_count):
                    self.doc.move_page(len(self.doc) - 1, insert_index + i)
            
            external_doc.close()
            return page_count
            
        except Exception as e:
            print(f"외부 PDF 병합 오류: {e}")
            return 0
    
    def on_drag_enter(self, event):
        """드래그 진입 시 시각적 피드백"""
        self.canvas.config(highlightcolor=self.selection_color, highlightthickness=2)
    
    def on_drag_leave(self, event):
        """드래그 이탈 시 시각적 피드백 제거"""
        self.canvas.config(highlightcolor="", highlightthickness=0)

    def on_window_resize(self, event):
        """창 크기가 변경될 때 썸네일 그리드 레이아웃 재조정"""
        # 창 크기 변경이 완료된 후 그리드 레이아웃 업데이트
        # 너무 자주 호출되지 않도록 디바운싱 적용
        if hasattr(self, '_resize_timer'):
            self.root.after_cancel(self._resize_timer)
        
        self._resize_timer = self.root.after(150, self.update_grid_layout)

if __name__ == "__main__":
    if TkinterDnD:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    root.geometry("1200x800")
    root.title("KUNHWA PDF Editor")
    
    app = PDFEditorApp(root)
    
    root.mainloop()
