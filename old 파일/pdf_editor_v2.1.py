import fitz
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, ttk
from PIL import Image, ImageTk
from functools import partial
import os
import sys

# 기본 텍스트 추출 기능만 사용

# tkinterdnd2 라이브러리 임포트 (윈도우 드래그 앤 드롭 지원)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False
    print("tkinterdnd2 라이브러리가 설치되지 않았습니다. 드래그 앤 드롭 기능을 사용할 수 없습니다.")
    print("설치 방법: pip install tkinterdnd2")

class ModernButton(tk.Button):
    """모던한 디자인의 버튼 클래스"""
    def __init__(self, parent, **kwargs):
        # 기본 스타일 설정
        default_style = {
            'font': ('맑은 고딕', 9, 'bold'),
            'relief': 'flat',
            'borderwidth': 0,
            'padx': 16,
            'pady': 6,
            'cursor': 'hand2',
            'activebackground': kwargs.get('bg', '#0078D4'),
            'activeforeground': 'white'
        }
        
        # 사용자 스타일과 기본 스타일 병합
        for key, value in default_style.items():
            if key not in kwargs:
                kwargs[key] = value
        
        super().__init__(parent, **kwargs)
        
        # 호버 효과
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)
        
        # 클릭 효과
        self.bind('<Button-1>', self.on_click)
        self.bind('<ButtonRelease-1>', self.on_release)
        
        self._original_bg = kwargs.get('bg', '#0078D4')
        self._pressed = False
    
    def on_enter(self, event):
        if not self._pressed:
            # 호버 시 약간 어둡게
            current_bg = self.cget('bg')
            darker_bg = self.darken_color(current_bg, 0.1)
            self.configure(bg=darker_bg)
    
    def on_leave(self, event):
        if not self._pressed:
            self.configure(bg=self._original_bg)
    
    def on_click(self, event):
        self._pressed = True
        # 클릭 시 더 어둡게
        darker_bg = self.darken_color(self._original_bg, 0.2)
        self.configure(bg=darker_bg)
    
    def on_release(self, event):
        self._pressed = False
        self.configure(bg=self._original_bg)
    
    def darken_color(self, color, factor):
        """색상을 어둡게 만드는 함수"""
        if color.startswith('#'):
            # 16진수 색상을 RGB로 변환
            r = int(color[1:3], 16)
            g = int(color[3:5], 16)
            b = int(color[5:7], 16)
            
            # 어둡게 만들기
            r = int(r * (1 - factor))
            g = int(g * (1 - factor))
            b = int(b * (1 - factor))
            
            # RGB를 16진수로 변환
            return f'#{r:02x}{g:02x}{b:02x}'
        return color

class PDFEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kunhwa PDF Editor")
        
        # 윈도우 스타일 설정
        self.root.configure(bg='#f8f9fa')
        
        # 상태 값들
        self.doc = None
        self.thumb_scale = 0.20  # 기본값 0.20
        self.preview_scale = 1.00  # 기본값 1.00
        self.current_page_index = 0  # 미리보기 표시할 페이지 인덱스
        self.selected_indices = set()
        self._zoom_target = 'thumbs'  # 줌 대상 패널
        
        self.thumbnails = []
        self.thumbnail_labels = []
        self.thumbnail_frames = []
        self.last_clicked_index = None
        self.drag_start_index = None
        self.drag_data = {"x": 0, "y": 0, "item": None}
        
        self.setup_ui()
        self.bind_events()

    def setup_drag_drop(self):
        """드래그 앤 드롭 설정"""
        if not DRAG_DROP_AVAILABLE:
            return
        
        try:
            # 메인 윈도우에 드래그 앤 드롭 바인딩
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_drop_file)
            
            # 썸네일 패널에도 드래그 앤 드롭 바인딩
            self.thumb_canvas.drop_target_register(DND_FILES)
            self.thumb_canvas.dnd_bind('<<Drop>>', self.on_drop_file)
            
            # 미리보기 패널에도 드래그 앤 드롭 바인딩
            self.preview_canvas.drop_target_register(DND_FILES)
            self.preview_canvas.dnd_bind('<<Drop>>', self.on_drop_file)
            
            print("드래그 앤 드롭 기능이 활성화되었습니다.")
        except Exception as e:
            print(f"드래그 앤 드롭 설정 중 오류: {e}")



    def open_pdf_from_path(self, file_path):
        """경로로부터 PDF 열기"""
        try:
            self.doc = fitz.open(file_path)
            self.refresh_thumbnails()
            self.update_preview()
            self.root.title(f"Kunhwa PDF Editor - {os.path.basename(file_path)}")
            print(f"PDF 파일을 열었습니다: {file_path}")
        except Exception as e:
            messagebox.showerror("오류", f"PDF 파일을 열 수 없습니다: {str(e)}")
            print(f"PDF 열기 오류: {e}")

    def create_pdf_from_image(self, image_path):
        """이미지로부터 새 PDF 생성"""
        try:
            # 이미지를 PDF로 변환
            img_doc = fitz.open()
            img_page = img_doc.new_page()
            
            # 이미지 삽입
            img_rect = fitz.Rect(0, 0, 595, 842)  # A4 크기
            img_page.insert_image(img_rect, filename=image_path)
            
            self.doc = img_doc
            self.refresh_thumbnails()
            self.update_preview()
            self.root.title(f"Kunhwa PDF Editor - {os.path.basename(image_path)}")
            print(f"이미지로부터 PDF를 생성했습니다: {image_path}")
        except Exception as e:
            messagebox.showerror("오류", f"이미지로부터 PDF를 생성할 수 없습니다: {str(e)}")
            print(f"이미지 PDF 생성 오류: {e}")

    def merge_pdf_from_path(self, file_path):
        """경로로부터 PDF 병합"""
        try:
            if not self.doc:
                messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
                return
            
            # PDF 파일 열기
            merge_doc = fitz.open(file_path)
            
            # 모든 페이지를 현재 문서에 추가
            for page_num in range(len(merge_doc)):
                page = merge_doc[page_num]
                self.doc.insert_pdf(merge_doc, from_page=page_num, to_page=page_num)
            
            merge_doc.close()
            self.refresh_thumbnails()
            self.update_preview()
            print(f"PDF 파일을 병합했습니다: {file_path}")
        except Exception as e:
            messagebox.showerror("오류", f"PDF 파일을 병합할 수 없습니다: {str(e)}")
            print(f"PDF 병합 오류: {e}")

    def merge_image_from_path(self, image_path):
        """경로로부터 이미지 병합"""
        try:
            if not self.doc:
                messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
                return
            
            # 새 페이지 생성
            page = self.doc.new_page()
            
            # 이미지 삽입
            img_rect = fitz.Rect(0, 0, 595, 842)  # A4 크기
            page.insert_image(img_rect, filename=image_path)
            
            self.refresh_thumbnails()
            self.update_preview()
            print(f"이미지를 병합했습니다: {image_path}")
        except Exception as e:
            messagebox.showerror("오류", f"이미지를 병합할 수 없습니다: {str(e)}")
            print(f"이미지 병합 오류: {e}")

    def on_drop_file(self, event):
        """파일 드롭 처리"""
        try:
            # 드롭된 파일 경로들 가져오기
            files = event.data
            
            # 윈도우 경로 형식 처리
            if files.startswith('{'):
                # 중괄호로 감싸진 경로들 처리
                files = files.strip('{}').split('} {')
            else:
                # 단일 파일 경로
                files = [files]
            
            # 지원하는 파일 형식 필터링
            supported_files = []
            for file_path in files:
                file_ext = os.path.splitext(file_path)[1].lower()
                if file_ext in ['.pdf', '.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif']:
                    supported_files.append(file_path)
            
            if not supported_files:
                messagebox.showwarning("경고", "지원하는 파일 형식: PDF, PNG, JPG, JPEG, BMP, TIFF, GIF")
                return
            
            # 드롭 위치에 따른 삽입 위치 결정
            if hasattr(self, 'drop_target_index'):
                # 썸네일 위에 직접 드롭된 경우
                drop_target = self.drop_target_index
                print(f"썸네일 위에 드롭됨: 타겟 위치 {drop_target}")
                # 사용 후 초기화
                delattr(self, 'drop_target_index')
            else:
                # 일반 드롭 처리 - 항상 마지막 위치에 병합
                drop_target = len(self.doc) if self.doc else 0
                print(f"일반 드롭: 마지막 위치로 설정 {drop_target}")
            
            print(f"최종 드롭 타겟 위치: {drop_target}")
            
            if not self.doc:
                # 현재 열린 PDF가 없으면 첫 번째 파일을 열기
                if os.path.splitext(supported_files[0])[1].lower() == '.pdf':
                    self.open_pdf_from_path(supported_files[0])
                    supported_files = supported_files[1:]  # 첫 번째는 이미 열었으므로 제외
                else:
                    # 첫 번째 파일이 이미지인 경우 새 PDF 생성
                    self.create_pdf_from_image(supported_files[0])
                    supported_files = supported_files[1:]  # 첫 번째는 이미 열었으므로 제외
                # 새로 열린 PDF의 경우 드롭 위치를 마지막으로 설정
                drop_target = len(self.doc) if self.doc else 0
                print(f"새 PDF 열기 후 마지막 위치로 설정: {drop_target}")
            
            # 나머지 파일들을 병합
            for file_path in supported_files:
                file_ext = os.path.splitext(file_path)[1].lower()
                if file_ext == '.pdf':
                    self.merge_pdf_from_path_with_position(file_path, drop_target)
                else:
                    self.merge_image_from_path_with_position(file_path, drop_target)
            
            # 드래그 앤 드롭으로 파일 병합 완료 시 메시지 표시하지 않음
                
        except Exception as e:
            messagebox.showerror("오류", f"파일 드롭 처리 중 오류가 발생했습니다: {str(e)}")
            print(f"드롭 오류 상세: {e}")

    def open_pdf_from_path(self, file_path):
        """경로로부터 PDF 열기"""
        try:
            self.doc = fitz.open(file_path)
            self.current_page_index = 0
            self.refresh_thumbnails()
            self.update_preview()
            print(f"PDF 열기 성공: {file_path}")
        except Exception as e:
            print(f"PDF 열기 실패: {e}")
            raise e

    def merge_pdf_from_path(self, file_path):
        """경로로부터 PDF 병합"""
        try:
            # 병합할 PDF 열기
            merge_doc = fitz.open(file_path)
            
            if not merge_doc:
                print(f"병합할 PDF 파일을 열 수 없습니다: {file_path}")
                return False
            
            # 삽입할 위치 결정
            if self.selected_indices:
                # 선택된 페이지 중 가장 앞쪽 위치에 삽입
                insert_pos = min(self.selected_indices)
            else:
                # 선택된 페이지가 없으면 맨 앞에 삽입
                insert_pos = 0
            
            # 병합할 PDF의 모든 페이지를 현재 문서에 추가
            added_pages = []
            for i in range(len(merge_doc)):
                try:
                    # 병합할 PDF의 페이지를 현재 문서에 복사
                    page = merge_doc[i]
                    
                    # 새 페이지 생성 (기존 페이지 크기 유지)
                    new_page = self.doc.new_page(width=page.rect.width, height=page.rect.height)
                    
                    # 페이지 내용 복사 (더 안전한 방법)
                    new_page.insert_image(new_page.rect, pixmap=page.get_pixmap())
                    
                    # 생성된 페이지를 임시로 저장
                    added_pages.append(len(self.doc) - 1)
                    
                except Exception as e:
                    print(f"페이지 {i} 복사 중 오류: {e}")
                    continue
            
            # 병합할 PDF 닫기
            merge_doc.close()
            
            if not added_pages:
                print(f"병합할 페이지가 없습니다: {file_path}")
                return False
            
            # 추가된 페이지들을 원하는 위치로 이동
            for i, page_index in enumerate(added_pages):
                try:
                    self.doc.move_page(page_index, insert_pos + i)
                except Exception as e:
                    print(f"페이지 이동 중 오류: {e}")
                    continue
            
            # 선택 상태 업데이트 (새로 추가된 페이지들 선택)
            new_selection = set(range(insert_pos, insert_pos + len(added_pages)))
            self.selected_indices = new_selection
            
            # 썸네일과 미리보기 새로고침
            self.refresh_thumbnails()
            self.update_preview()
            
            print(f"PDF 병합 성공: {file_path}, {len(added_pages)}개 페이지 추가")
            return True
            
        except Exception as e:
            print(f"PDF 병합 실패: {e}")
            return False

    def get_drop_target_from_event(self, event):
        """드롭 이벤트에서 타겟 위치 계산"""
        try:
            # 드롭된 위젯 확인
            widget = event.widget
            print(f"드롭된 위젯: {widget}")
            
            if widget == self.thumb_canvas:
                # 썸네일 캔버스에 드롭된 경우
                print(f"썸네일 캔버스에 드롭됨: x={event.x}, y={event.y}")
                drop_pos = self.get_thumbnail_drop_position(event)
                print(f"썸네일 드롭 위치 계산 결과: {drop_pos}")
                return drop_pos
            elif widget == self.preview_canvas:
                # 미리보기 캔버스에 드롭된 경우
                print(f"미리보기 캔버스에 드롭됨: 현재 페이지={self.current_page_index}")
                return self.current_page_index
            else:
                # 메인 윈도우에 드롭된 경우 - 마지막 위치에 병합
                print(f"메인 윈도우에 드롭됨: 마지막 위치로 설정")
                return -1  # 마지막 위치를 나타내는 특별한 값
        except Exception as e:
            print(f"드롭 타겟 계산 오류: {e}")
            return 0

    def get_thumbnail_drop_position(self, event):
        """썸네일 영역에서 드롭 위치 계산"""
        try:
            # 마우스 위치를 캔버스 좌표로 변환
            canvas_x = self.thumb_canvas.canvasx(event.x)
            canvas_y = self.thumb_canvas.canvasy(event.y)
            
            # 스크롤 위치 고려
            scroll_y = self.thumb_canvas.yview()[0] * self.thumb_scrollable_frame.winfo_height()
            adjusted_y = canvas_y + scroll_y
            
            print(f"드롭 좌표: canvas_x={canvas_x}, canvas_y={canvas_y}, adjusted_y={adjusted_y}")
            
            if self.thumbnail_frames:
                # 그리드 레이아웃 계산
                sample_width = self.thumbnail_frames[0].winfo_reqwidth() + 10
                sample_height = self.thumbnail_frames[0].winfo_reqheight() + 10
                canvas_width = self.thumb_canvas.winfo_width()
                col_count = max(canvas_width // sample_width, 1)
                
                print(f"레이아웃 정보: sample_width={sample_width}, sample_height={sample_height}, col_count={col_count}")
                
                # 행과 열 계산
                row = int(adjusted_y // sample_height)
                col = int(canvas_x // sample_width)
                
                # 인덱스 계산
                index = row * col_count + col
                
                # 마지막 썸네일의 위치 계산
                last_row = (len(self.thumbnail_frames) - 1) // col_count
                last_thumbnail_y = last_row * sample_height
                
                print(f"계산 결과: row={row}, col={col}, index={index}, last_row={last_row}, last_thumbnail_y={last_thumbnail_y}")
                
                # 유효한 범위인지 확인
                if 0 <= index < len(self.thumbnail_frames):
                    print(f"드롭 위치 계산: 행={row}, 열={col}, 인덱스={index}")
                    return index
                elif index >= len(self.thumbnail_frames) or adjusted_y > last_thumbnail_y + sample_height:
                    # 마지막 썸네일 뒤나 아래 빈공간에 드롭된 경우
                    print(f"드롭 위치 계산: 마지막 위치 {len(self.thumbnail_frames)} (빈공간 드롭)")
                    return -1  # 마지막 위치를 나타내는 특별한 값
                else:
                    print(f"드롭 위치 계산: 기본 위치를 마지막으로 설정")
                    return -1  # 기본 위치도 마지막으로 설정
            
            # 썸네일이 없는 경우도 마지막 위치로 처리
            print("썸네일이 없음: 마지막 위치로 설정")
            return -1
        except Exception as e:
            print(f"썸네일 드롭 위치 계산 오류: {e}")
            return -1  # 오류 발생 시에도 마지막 위치로 설정

    def merge_pdf_from_path_with_position(self, file_path, insert_pos):
        """지정된 위치에 PDF 병합"""
        try:
            # 병합할 PDF 열기
            merge_doc = fitz.open(file_path)
            
            if not merge_doc:
                print(f"병합할 PDF 파일을 열 수 없습니다: {file_path}")
                return False
            
            # 병합할 PDF의 모든 페이지를 현재 문서에 추가
            added_pages = []
            for i in range(len(merge_doc)):
                try:
                    # 병합할 PDF의 페이지를 현재 문서에 복사
                    page = merge_doc[i]
                    
                    # 새 페이지 생성 (기존 페이지 크기 유지)
                    new_page = self.doc.new_page(width=page.rect.width, height=page.rect.height)
                    
                    # 페이지 내용 복사 (더 안전한 방법)
                    new_page.insert_image(new_page.rect, pixmap=page.get_pixmap())
                    
                    # 생성된 페이지를 임시로 저장
                    added_pages.append(len(self.doc) - 1)
                    
                except Exception as e:
                    print(f"페이지 {i} 복사 중 오류: {e}")
                    continue
            
            # 병합할 PDF 닫기
            merge_doc.close()
            
            if not added_pages:
                print(f"병합할 페이지가 없습니다: {file_path}")
                return False
            
            # 추가된 페이지들을 원하는 위치로 이동
            for i, page_index in enumerate(added_pages):
                try:
                    self.doc.move_page(page_index, insert_pos + i)
                except Exception as e:
                    print(f"페이지 이동 중 오류: {e}")
                    continue
            
            # 선택 상태 업데이트 (새로 추가된 페이지들 선택)
            new_selection = set(range(insert_pos, insert_pos + len(added_pages)))
            self.selected_indices = new_selection
            
            # 썸네일과 미리보기 새로고침
            self.refresh_thumbnails()
            self.update_preview()
            
            print(f"PDF 병합 성공: {file_path}, {len(added_pages)}개 페이지을 위치 {insert_pos}에 추가")
            return True
            
        except Exception as e:
            print(f"PDF 병합 실패: {e}")
            return False

    def create_pdf_from_image(self, image_path):
        """이미지로부터 새 PDF 생성"""
        try:
            # 이미지 열기
            img = Image.open(image_path)
            
            # PDF 문서 생성
            self.doc = fitz.open()
            
            # 이미지 크기를 A4 크기로 조정 (비율 유지)
            a4_width = 595.276  # A4 가로 (포인트)
            a4_height = 841.890  # A4 세로 (포인트)
            
            # 이미지 비율 계산
            img_ratio = img.width / img.height
            a4_ratio = a4_width / a4_height
            
            if img_ratio > a4_ratio:
                # 이미지가 가로로 긴 경우
                new_width = a4_width
                new_height = a4_width / img_ratio
            else:
                # 이미지가 세로로 긴 경우
                new_height = a4_height
                new_width = a4_height * img_ratio
            
            # 새 페이지 생성
            page = self.doc.new_page(width=new_width, height=new_height)
            
            # 이미지를 PDF 페이지에 삽입
            page.insert_image(page.rect, filename=image_path)
            
            self.current_page_index = 0
            self.refresh_thumbnails()
            self.update_preview()
            
            print(f"이미지로부터 PDF 생성 성공: {image_path}")
            
        except Exception as e:
            print(f"이미지로부터 PDF 생성 실패: {e}")
            raise e

    def merge_image_from_path_with_position(self, image_path, insert_pos):
        """지정된 위치에 이미지 병합"""
        try:
            # 이미지 열기
            img = Image.open(image_path)
            
            # 이미지 크기를 A4 크기로 조정 (비율 유지)
            a4_width = 595.276  # A4 가로 (포인트)
            a4_height = 841.890  # A4 세로 (포인트)
            
            # 이미지 비율 계산
            img_ratio = img.width / img.height
            a4_ratio = a4_width / a4_height
            
            if img_ratio > a4_ratio:
                # 이미지가 가로로 긴 경우
                new_width = a4_width
                new_height = a4_width / img_ratio
            else:
                # 이미지가 세로로 긴 경우
                new_height = a4_height
                new_width = a4_height * img_ratio
            
            # 새 페이지 생성
            new_page = self.doc.new_page(width=new_width, height=new_height)
            
            # 이미지를 PDF 페이지에 삽입
            new_page.insert_image(new_page.rect, filename=image_path)
            
            # 생성된 페이지를 원하는 위치로 이동
            try:
                self.doc.move_page(len(self.doc) - 1, insert_pos)
            except Exception as e:
                print(f"페이지 이동 중 오류: {e}")
            
            # 선택 상태 업데이트 (새로 추가된 페이지 선택)
            self.selected_indices = {insert_pos}
            
            # 썸네일과 미리보기 새로고침
            self.refresh_thumbnails()
            self.update_preview()
            
            print(f"이미지 병합 성공: {image_path}, 위치 {insert_pos}에 추가")
            return True
            
        except Exception as e:
            print(f"이미지 병합 실패: {e}")
            return False

    def setup_ui(self):
        # 상단 버튼 프레임 (그림자 효과를 위한 컨테이너)
        top_container = tk.Frame(self.root, bg='#e9ecef', height=80)
        top_container.pack(side=tk.TOP, fill=tk.X)
        top_container.pack_propagate(False)
        
        # 상단 버튼 프레임
        top_frame = tk.Frame(top_container, bg='#ffffff', relief='flat', bd=0)
        top_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        # PDF 작업 버튼들 (최적화된 색상과 크기)
        pdf_frame = tk.Frame(top_frame, bg='#ffffff')
        pdf_frame.pack(side=tk.LEFT, padx=8)
        
        ModernButton(pdf_frame, text="PDF 열기", command=self.open_pdf, 
                    bg="#2563EB", fg="white").pack(side=tk.LEFT, padx=4)
        ModernButton(pdf_frame, text="PDF 병합", command=self.merge_pdf, 
                    bg="#3B82F6", fg="white").pack(side=tk.LEFT, padx=4)
        ModernButton(pdf_frame, text="PDF 저장", command=self.save_pdf, 
                    bg="#1D4ED8", fg="white").pack(side=tk.LEFT, padx=4)
        
        # 구분선 추가 (모던한 스타일)
        separator1 = tk.Frame(top_frame, width=1, height=36, bg="#E5E7EB")
        separator1.pack(side=tk.LEFT, padx=16)
        
        # 회전 버튼들 (최적화된 색상과 크기)
        rotate_frame = tk.Frame(top_frame, bg='#ffffff')
        rotate_frame.pack(side=tk.LEFT, padx=8)
        
        ModernButton(rotate_frame, text="↱ 우로90°회전", command=lambda: self.rotate_pages(90), 
                    bg="#F59E0B", fg="white").pack(side=tk.LEFT, padx=4)
        ModernButton(rotate_frame, text="↰ 좌로90°회전", command=lambda: self.rotate_pages(-90), 
                    bg="#D97706", fg="white").pack(side=tk.LEFT, padx=4)
        
        # 구분선 추가
        separator2 = tk.Frame(top_frame, width=1, height=36, bg="#E5E7EB")
        separator2.pack(side=tk.LEFT, padx=16)
        
        # 빈페이지 삽입 버튼 (최적화된 색상과 크기)
        ModernButton(top_frame, text="빈페이지 삽입", command=self.show_insert_blank_page_dialog, 
                    bg="#10B981", fg="white").pack(side=tk.LEFT, padx=8)
        
        # 구분선 추가
        separator3 = tk.Frame(top_frame, width=1, height=36, bg="#E5E7EB")
        separator3.pack(side=tk.LEFT, padx=16)
        
        # 텍스트 추출 버튼 (최적화된 색상과 크기)
        text_frame = tk.Frame(top_frame, bg='#ffffff')
        text_frame.pack(side=tk.LEFT, padx=8)
        
        ModernButton(text_frame, text="텍스트 추출", command=self.extract_text_directly, 
                    bg="#F59E0B", fg="white").pack(side=tk.LEFT, padx=4)
        
        # 페이지 맞춤 버튼을 제일 오른쪽에 배치 (최적화된 색상과 크기)
        ModernButton(top_frame, text="페이지에 맞춤", command=self.fit_page_to_screen, 
                    bg="#6366F1", fg="white").pack(side=tk.RIGHT, padx=8)
        
        # 정보 표시 프레임 (버튼 아래) - 모던한 카드 스타일
        info_container = tk.Frame(self.root, bg='#f8f9fa', height=50)
        info_container.pack(side=tk.TOP, fill=tk.X)
        info_container.pack_propagate(False)
        
        info_frame = tk.Frame(info_container, bg="white", relief="flat", bd=0)
        info_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 선택된 페이지 정보 (왼쪽) - 더 진한 폰트
        self.selection_info = tk.Label(info_frame, text="선택된 페이지: 없음", 
                                     bg="white", fg="#212529", font=("맑은 고딕", 11, "bold"))
        self.selection_info.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 키보드 단축키 안내 (오른쪽) - 더 진한 폰트
        shortcuts_info = tk.Label(info_frame, 
                                text="Ctrl+클릭: 다중선택 | Shift+클릭: 범위선택 | Delete: 삭제 | Ctrl+A: 전체선택 | ESC: 선택해제",
                                bg="white", fg="#495057", font=("맑은 고딕", 9, "bold"))
        shortcuts_info.pack(side=tk.RIGHT, padx=10, pady=5)
        
        # 수평 분할 레이아웃 (PanedWindow) - 모던한 스타일
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 좌측 패널: 썸네일 목록
        self.setup_thumbnail_panel()
        
        # 우측 패널: 미리보기
        self.setup_preview_panel()
        
        # 패널 크기 설정
        self.paned_window.sashpos(0, 320)  # 좌측 최소 폭 320px
        
        # 하단 저작권 정보 프레임 - 모던한 스타일
        copyright_container = tk.Frame(self.root, bg='#e9ecef', height=35)
        copyright_container.pack(side=tk.BOTTOM, fill=tk.X)
        copyright_container.pack_propagate(False)
        
        copyright_frame = tk.Frame(copyright_container, bg="white", relief="flat", bd=0)
        copyright_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=8)
        
        # 저작권 문구 (중앙 정렬) - 더 진한 폰트
        copyright_label = tk.Label(copyright_frame, 
                                 text="© 2025 Kunhwa Engineering & Consulting | Developed by TaeWoong Jang",
                                 bg="white", fg="#495057", font=("맑은 고딕", 8, "bold"))
        copyright_label.pack(expand=True, pady=2)
        
        # 드래그 앤 드롭 기능 설정
        self.setup_drag_drop()

    def setup_thumbnail_panel(self):
        # 좌측 썸네일 패널 - 모던한 스타일
        left_frame = tk.Frame(self.paned_window, bg='#ffffff', relief='flat', bd=0)
        self.paned_window.add(left_frame, weight=1)
        
        # 패널 제목
        title_frame = tk.Frame(left_frame, bg='#f8f9fa', height=40)
        title_frame.pack(side=tk.TOP, fill=tk.X)
        title_frame.pack_propagate(False)
        
        # 패널 제목 - 더 진한 폰트
        title_label = tk.Label(title_frame, text="페이지 썸네일", 
                               bg='#f8f9fa', fg='#212529', font=("맑은 고딕", 12, "bold"))
        title_label.pack(expand=True)
        
        # 썸네일 캔버스와 스크롤바 - 모던한 스타일
        self.thumb_canvas = tk.Canvas(left_frame, bg="white", highlightthickness=0, relief="flat")
        self.thumb_scrollbar = tk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.thumb_canvas.yview)
        self.thumb_scrollable_frame = tk.Frame(self.thumb_canvas, bg="white")
        
        self.thumb_canvas.create_window((0, 0), window=self.thumb_scrollable_frame, anchor="nw")
        self.thumb_canvas.configure(yscrollcommand=self.thumb_scrollbar.set)
        
        self.thumb_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.thumb_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        # 썸네일 패널 이벤트 바인딩
        self.thumb_canvas.bind("<Configure>", self.schedule_grid_update)
        self.thumb_canvas.bind("<Enter>", lambda e: self.set_zoom_target('thumbs'))
        self.thumb_canvas.bind("<MouseWheel>", self.on_thumb_mousewheel)
        
        # 썸네일 프레임에도 마우스 휠 이벤트 바인딩
        self.thumb_scrollable_frame.bind("<MouseWheel>", self.on_thumb_mousewheel)
        
        # 패널 크기 변경 시 레이아웃 업데이트
        self.paned_window.bind("<Configure>", self.on_paned_configure)

    def setup_preview_panel(self):
        # 우측 미리보기 패널 - 모던한 스타일
        right_frame = tk.Frame(self.paned_window, bg='#ffffff', relief='flat', bd=0)
        self.paned_window.add(right_frame, weight=2)
        
        # 패널 제목
        title_frame = tk.Frame(right_frame, bg='#f8f9fa', height=40)
        title_frame.pack(side=tk.TOP, fill=tk.X)
        title_frame.pack_propagate(False)
        
        # 패널 제목 - 더 진한 폰트
        title_label = tk.Label(title_frame, text="페이지 미리보기", 
                               bg='#f8f9fa', fg='#212529', font=("맑은 고딕", 12, "bold"))
        title_label.pack(expand=True)
        
        # 미리보기 캔버스와 스크롤바 - 모던한 스타일
        self.preview_canvas = tk.Canvas(right_frame, bg="white", highlightthickness=0, relief="flat")
        self.preview_v_scrollbar = tk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        self.preview_h_scrollbar = tk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=self.preview_canvas.xview)
        
        self.preview_canvas.configure(yscrollcommand=self.preview_v_scrollbar.set, xscrollcommand=self.preview_h_scrollbar.set)
        
        # 스크롤바 배치
        self.preview_v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        self.preview_h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X, padx=5)
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 미리보기 패널 이벤트 바인딩
        self.preview_canvas.bind("<Enter>", lambda e: self.set_zoom_target('preview'))
        self.preview_canvas.bind("<MouseWheel>", self.on_preview_mousewheel)
        self.preview_canvas.bind("<Shift-MouseWheel>", self.on_preview_shift_mousewheel)
        self.preview_canvas.bind("<Configure>", self.on_preview_configure)
        
        # 로고 표시 (PDF가 로드되지 않았을 때)
        self.show_logo()

    def show_logo(self):
        """로고 표시"""
        try:
            # 로고 파일 경로 (현재 코드 파일과 같은 디렉토리의 data 폴더)
            current_dir = os.path.dirname(os.path.abspath(__file__))
            logo_path = os.path.join(current_dir, "data", "kunhwa_logo.png")
            
            # 로고 파일이 존재하는지 확인
            if os.path.exists(logo_path):
                # 로고 이미지 로드
                logo_img = Image.open(logo_path)
                
                # 캔버스 크기에 맞게 리사이즈
                canvas_width = self.preview_canvas.winfo_width()
                canvas_height = self.preview_canvas.winfo_height()
                
                if canvas_width > 1 and canvas_height > 1:
                    # 로고 크기 계산 (캔버스의 40% 크기로 제한)
                    max_logo_width = int(canvas_width * 0.4)
                    max_logo_height = int(canvas_height * 0.4)
                    
                    # 비율 유지하면서 리사이즈
                    logo_img.thumbnail((max_logo_width, max_logo_height), Image.Resampling.LANCZOS)
                    
                    # PhotoImage로 변환
                    self.logo_photo = ImageTk.PhotoImage(logo_img)
                    
                    # 캔버스에 로고 배치 (중앙)
                    logo_x = (canvas_width - logo_img.width) // 2
                    logo_y = (canvas_height - logo_img.height) // 2
                    
                    self.preview_canvas.create_image(logo_x, logo_y, anchor="nw", image=self.logo_photo, tags="logo")
                    
                    print(f"Kunhwa 로고가 성공적으로 표시되었습니다. 경로: {logo_path}")
                else:
                    # 캔버스가 아직 렌더링되지 않았으면 나중에 다시 시도
                    self.root.after(100, self.show_logo)
            else:
                # 로고 파일이 없으면 오류 메시지
                print(f"로고 파일을 찾을 수 없습니다: {logo_path}")
                self.show_text_logo()
                
        except Exception as e:
            print(f"로고 표시 중 오류: {e}")
            # 오류 발생 시 텍스트 로고 표시
            self.show_text_logo()

    def show_text_logo(self):
        """텍스트 로고 표시 (이미지 로고를 불러올 수 없을 때)"""
        try:
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:
                # 텍스트 로고 생성
                logo_text = "Kunhwa PDF Editor"
                
                # 캔버스에 텍스트 로고 배치 (중앙)
                self.preview_canvas.create_text(
                    canvas_width // 2, 
                    canvas_height // 2, 
                    text=logo_text, 
                    font=("맑은 고딕", 24, "bold"), 
                    fill="gray",
                    tags="logo"
                )
                
                print("텍스트 로고가 표시되었습니다.")
            else:
                # 캔버스가 아직 렌더링되지 않았으면 나중에 다시 시도
                self.root.after(100, self.show_text_logo)
                
        except Exception as e:
            print(f"텍스트 로고 표시 중 오류: {e}")

    def bind_events(self):
        # 전역 이벤트 바인딩
        self.root.bind_all("<Control-MouseWheel>", self.on_ctrl_mousewheel)
        # Delete 키로 페이지 삭제
        self.root.bind_all("<Delete>", self.delete_pages_with_key)
        # ESC 키로 다중선택 해제
        self.root.bind_all("<Escape>", self.clear_selection)
        # Ctrl+A로 전체 선택
        self.root.bind_all("<Control-a>", self.select_all_pages)
        self.root.bind_all("<Control-A>", self.select_all_pages)

    def set_zoom_target(self, target):
        """줌 대상 패널 설정"""
        self._zoom_target = target

    def on_ctrl_mousewheel(self, event):
        """Ctrl + 마우스 휠로 줌 인/아웃"""
        if self._zoom_target == 'thumbs':
            # 썸네일 줌
            if event.delta > 0:
                self.thumb_scale = min(1.50, self.thumb_scale + 0.05)
            else:
                self.thumb_scale = max(0.05, self.thumb_scale - 0.05)
            self.refresh_thumbnails()
        elif self._zoom_target == 'preview':
            # 미리보기 줌
            if event.delta > 0:
                self.preview_scale = min(4.00, self.preview_scale + 0.05)
            else:
                self.preview_scale = max(0.10, self.preview_scale - 0.05)
            self.update_preview()

    def on_thumb_mousewheel(self, event):
        """썸네일 패널 마우스 휠 스크롤"""
        self.thumb_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def on_preview_mousewheel(self, event):
        """미리보기 패널 마우스 휠 스크롤 (세로)"""
        # Ctrl 키가 눌려있으면 페이지 전환 없이 일반 스크롤만
        if event.state & 0x0004:  # Ctrl 키 상태 확인
            self.preview_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            return
        
        # 전체 페이지가 화면에 다 보이는지 확인
        if self.is_page_fit_to_screen():
            # 전체 페이지가 보이는 경우, 페이지 이동
            if event.delta > 0:
                # 위로 스크롤 - 이전 페이지
                self.go_to_previous_page()
            else:
                # 아래로 스크롤 - 다음 페이지
                self.go_to_next_page()
        else:
            # 페이지가 화면보다 큰 경우, 스크롤 후 끝에 도달하면 페이지 변경
            self.scroll_with_page_change(event)

    def scroll_with_page_change(self, event):
        """스크롤 후 끝에 도달하면 페이지 변경"""
        try:
            # Ctrl 키가 눌려있으면 페이지 전환 없이 일반 스크롤만
            if event.state & 0x0004:  # Ctrl 키 상태 확인
                self.preview_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                return
            
            # 현재 스크롤 위치 가져오기
            current_scroll = self.preview_canvas.yview()
            scroll_top = current_scroll[0]
            scroll_bottom = current_scroll[1]
            
            # 스크롤 방향에 따른 처리
            if event.delta > 0:  # 위로 스크롤 (휠을 위로)
                # 현재 스크롤이 맨 위에 있는지 확인
                if scroll_top <= 0.01:  # 맨 위에 거의 도달
                    # 이전 페이지로 이동
                    self.go_to_previous_page()
                else:
                    # 일반 스크롤
                    self.preview_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            else:  # 아래로 스크롤 (휠을 아래로)
                # 현재 스크롤이 맨 아래에 있는지 확인
                if scroll_bottom >= 0.99:  # 맨 아래에 거의 도달
                    # 다음 페이지로 이동
                    self.go_to_next_page()
                else:
                    # 일반 스크롤
                    self.preview_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                    
        except Exception as e:
            print(f"스크롤 페이지 변경 중 오류: {e}")
            # 오류 발생 시 일반 스크롤로 폴백
            self.preview_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def is_page_fit_to_screen(self):
        """페이지가 화면에 완전히 맞는지 확인"""
        if not self.doc or self.current_page_index >= len(self.doc):
            return False
        
        try:
            # 현재 페이지의 크기 가져오기
            page = self.doc[self.current_page_index]
            page_width = page.rect.width
            page_height = page.rect.height
            
            # 미리보기 패널의 크기 가져오기
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            if canvas_width <= 1 or canvas_height <= 1:
                return False
            
            # 페이지가 화면에 완전히 맞는지 확인 (여백 20px 고려)
            return (page_width * self.preview_scale <= canvas_width - 20 and 
                    page_height * self.preview_scale <= canvas_height - 20)
            
        except Exception as e:
            print(f"페이지 맞춤 확인 중 오류: {e}")
            return False

    def go_to_previous_page(self):
        """이전 페이지로 이동"""
        if self.doc and self.current_page_index > 0:
            self.current_page_index -= 1
            self.update_preview()
            # 스크롤 위치를 맨 위로 초기화
            self.preview_canvas.yview_moveto(0)
            # 썸네일에서도 해당 페이지 선택
            self.selected_indices = {self.current_page_index}
            self.update_selection_highlight()
            # 썸네일 패널에서 해당 페이지가 보이도록 스크롤
            self.scroll_to_thumbnail(self.current_page_index)

    def go_to_next_page(self):
        """다음 페이지로 이동"""
        if self.doc and self.current_page_index < len(self.doc) - 1:
            self.current_page_index += 1
            self.update_preview()
            # 스크롤 위치를 맨 위로 초기화
            self.preview_canvas.yview_moveto(0)
            # 썸네일에서도 해당 페이지 선택
            self.selected_indices = {self.current_page_index}
            self.update_selection_highlight()
            # 썸네일 패널에서 해당 페이지가 보이도록 스크롤
            self.scroll_to_thumbnail(self.current_page_index)

    def scroll_to_thumbnail(self, page_index):
        """썸네일 패널에서 특정 페이지가 보이도록 스크롤"""
        try:
            if not self.thumbnail_frames or page_index >= len(self.thumbnail_frames):
                return
            
            # 해당 썸네일 프레임의 위치 계산
            target_frame = self.thumbnail_frames[page_index]
            target_frame.update_idletasks()
            
            # 프레임의 y 위치 계산
            frame_y = target_frame.winfo_y()
            
            # 스크롤 위치 조정
            if frame_y > 0:
                # 프레임이 아래쪽에 있는 경우, 위로 스크롤
                scroll_ratio = max(0, (frame_y - 100) / self.thumb_scrollable_frame.winfo_height())
                self.thumb_canvas.yview_moveto(scroll_ratio)
            elif frame_y < 0:
                # 프레임이 위쪽에 있는 경우, 아래로 스크롤
                scroll_ratio = max(0, (frame_y + 100) / self.thumb_scrollable_frame.winfo_height())
                self.thumb_canvas.yview_moveto(scroll_ratio)
                
        except Exception as e:
            print(f"썸네일 스크롤 중 오류: {e}")

    def on_preview_shift_mousewheel(self, event):
        """미리보기 패널 Shift + 마우스 휠 스크롤 (가로)"""
        self.preview_canvas.xview_scroll(int(-1*(event.delta/120)), "units")

    def on_preview_configure(self, event):
        """미리보기 패널 크기 조정 시 이미지 재배치"""
        if hasattr(self, 'preview_image') and self.preview_image:
            self.update_preview()
        else:
            # PDF가 로드되지 않았으면 로고 재배치
            self.show_logo()

    def on_paned_configure(self, event):
        """패널 크기 변경 시 썸네일 레이아웃 재계산"""
        self.schedule_grid_update()
        
        # 창 크기 변경 후 스크롤 영역도 확인
        if hasattr(self, 'thumbnail_frames') and self.thumbnail_frames:
            self.root.after(200, self.ensure_scroll_region)

    def open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not path:
            return
        self.doc = fitz.open(path)
        self.current_page_index = 0
        self.refresh_thumbnails()
        self.update_preview()

    def refresh_thumbnails(self):
        if not self.doc:
            return
        
        for widget in self.thumb_scrollable_frame.winfo_children():
            widget.destroy()
        
        self.thumbnails.clear()
        self.thumbnail_labels.clear()
        self.thumbnail_frames = []
        
        for i in range(len(self.doc)):
            pix = self.doc[i].get_pixmap(matrix=fitz.Matrix(self.thumb_scale, self.thumb_scale))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            thumb = ImageTk.PhotoImage(img)
            self.thumbnails.append(thumb)
            
            # 모던한 썸네일 프레임 스타일
            frame = tk.Frame(self.thumb_scrollable_frame, bg="white", relief="flat", bd=0)
            
            # 썸네일 이미지 라벨 - 모던한 스타일
            label = tk.Label(frame, image=thumb, borderwidth=2, relief="solid", bg="white")
            label.bind("<Button-1>", partial(self.handle_selection, index=i))
            label.bind("<Double-Button-1>", partial(self.on_double_click, index=i))
            label.bind("<B1-Motion>", partial(self.on_drag, index=i))
            label.bind("<ButtonRelease-1>", partial(self.on_drop, index=i))
            label.bind("<Enter>", lambda e, idx=i: self.on_enter(e, idx))
            label.bind("<Leave>", lambda e, idx=i: self.on_leave(e, idx))
            # 각 썸네일 라벨에도 마우스 휠 이벤트 바인딩
            label.bind("<MouseWheel>", self.on_thumb_mousewheel)
            
            # 각 썸네일 라벨에 드래그 앤 드롭 이벤트 바인딩
            if DRAG_DROP_AVAILABLE:
                try:
                    label.drop_target_register(DND_FILES)
                    label.dnd_bind('<<Drop>>', partial(self.on_drop_on_thumbnail, index=i))
                except:
                    pass
            
            label.pack(padx=5, pady=5)
            
            # 페이지 번호 라벨 추가 - 더 진한 폰트
            num_label = tk.Label(frame, text=f"{i+1}", font=("맑은 고딕", 10, "bold"), 
                               bg="white", fg="#212529")
            num_label.pack(pady=(0, 5))
            # 페이지 번호 라벨에도 마우스 휠 이벤트 바인딩
            num_label.bind("<MouseWheel>", self.on_thumb_mousewheel)
            
            # 페이지 번호 라벨에도 드래그 앤 드롭 이벤트 바인딩
            if DRAG_DROP_AVAILABLE:
                try:
                    num_label.drop_target_register(DND_FILES)
                    num_label.dnd_bind('<<Drop>>', partial(self.on_drop_on_thumbnail, index=i))
                except:
                    pass
            
            self.thumbnail_labels.append(label)
            self.thumbnail_frames.append(frame)
        
        # 썸네일 생성 후 레이아웃 업데이트
        self.root.after(50, self.update_grid_layout)  # 더 긴 지연으로 안정성 향상
        
        # 스크롤 영역이 제대로 설정되었는지 확인
        self.root.after(100, self.ensure_scroll_region)
        
        self.update_selection_highlight()

    def on_drop_on_thumbnail(self, event, index):
        """특정 썸네일 위에 파일 드롭 처리"""
        print(f"썸네일 {index} 위에 파일 드롭됨")
        # 해당 썸네일 앞에 병합하도록 설정
        self.drop_target_index = index
        # 일반 드롭 처리 함수 호출
        self.on_drop_file(event)

    def update_preview(self):
        """미리보기 패널 업데이트"""
        if not self.doc or self.current_page_index >= len(self.doc):
            # PDF가 없으면 로고 표시
            self.show_logo()
            return
        
        # 기존 이미지와 로고 제거
        self.preview_canvas.delete("all")
        
        # 현재 페이지 렌더링
        page = self.doc[self.current_page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(self.preview_scale, self.preview_scale))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        preview_img = ImageTk.PhotoImage(img)
        
        # 이미지 참조 유지
        self.preview_image = preview_img
        
        # 캔버스에 이미지 배치 (중앙 정렬)
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()
        
        if canvas_width > 1 and canvas_height > 1:  # 캔버스가 실제로 렌더링된 후
            x = max(0, (canvas_width - pix.width) // 2)
            y = max(0, (canvas_height - pix.height) // 2)
            self.preview_canvas.create_image(x, y, anchor="nw", image=preview_img)
            
            # 스크롤 영역 설정
            self.preview_canvas.config(scrollregion=(0, 0, pix.width, pix.height))

    def schedule_grid_update(self, event=None):
        self.root.after(10, self.update_grid_layout)
        
        # 그리드 업데이트 후 스크롤 영역도 확인
        if hasattr(self, 'thumbnail_frames') and self.thumbnail_frames:
            self.root.after(100, self.ensure_scroll_region)

    def update_grid_layout(self):
        """썸네일을 창 크기에 맞춰 자동으로 배열하는 그리드 레이아웃"""
        if not self.doc or not self.thumbnail_frames:
            return

        # 캔버스 크기 가져오기
        canvas_width = self.thumb_canvas.winfo_width()
        canvas_height = self.thumb_canvas.winfo_height()
        
        if canvas_width <= 1 or canvas_height <= 1:
            return  # 캔버스가 아직 렌더링되지 않음
        
        # 썸네일 프레임의 실제 크기 계산
        if self.thumbnail_frames:
            sample_frame = self.thumbnail_frames[0]
            sample_frame.update_idletasks()  # 실제 크기 계산을 위해 업데이트
            
            # 패딩 포함한 썸네일 크기
            thumb_width = sample_frame.winfo_reqwidth() + 10  # 좌우 패딩 5px씩
            thumb_height = sample_frame.winfo_reqheight() + 10  # 상하 패딩 5px씩
            
            # 사용 가능한 열 수 계산 (최소 1열 보장)
            available_width = canvas_width - 20  # 좌우 여백 10px씩
            col_count = max(1, available_width // thumb_width)
            
            # 행 수 계산
            total_pages = len(self.thumbnail_frames)
            row_count = (total_pages + col_count - 1) // col_count  # 올림 나눗셈
            
            # 그리드 배치
            for i, frame in enumerate(self.thumbnail_frames):
                row = i // col_count
                col = i % col_count
                
                # 그리드 위치 설정
                frame.grid(row=row, column=col, padx=5, pady=5, sticky="nw")
                
                # 프레임이 잘리지 않도록 확인
                frame.grid_propagate(False)
            
            # 스크롤 영역 갱신
            self.thumb_canvas.update_idletasks()
            
            # 전체 그리드 크기 계산
            total_width = col_count * thumb_width
            total_height = row_count * thumb_height
            
            # 스크롤 영역 설정 (여백 포함) - 하단에 2열 정도 여유 공간 추가
            # 2열의 높이만큼 여유 공간 계산
            extra_height = thumb_height * 2 + 20  # 2열 + 패딩
            
            # 스크롤 영역 설정 (하단 여유 공간 포함)
            self.thumb_canvas.config(scrollregion=(0, 0, total_width, total_height + extra_height))

    def ensure_scroll_region(self):
        """스크롤 영역이 제대로 설정되었는지 확인하고 필요시 수정"""
        if not self.doc or not self.thumbnail_frames:
            return
            
        try:
            # 현재 스크롤 영역 가져오기
            current_scroll = self.thumb_canvas.bbox("all")
            if not current_scroll:
                return
                
            # 캔버스 크기
            canvas_width = self.thumb_canvas.winfo_width()
            canvas_height = self.thumb_canvas.winfo_height()
            
            if canvas_width <= 1 or canvas_height <= 1:
                return
                
            # 썸네일 프레임의 실제 크기 계산
            if self.thumbnail_frames:
                sample_frame = self.thumbnail_frames[0]
                sample_frame.update_idletasks()
                thumb_height = sample_frame.winfo_reqheight() + 10
                
                # 2열의 높이만큼 여유 공간 계산
                extra_height = thumb_height * 2 + 20
                
                # 스크롤 영역 계산 - 하단에 2열 여유 공간 추가
                scroll_width = max(current_scroll[2], canvas_width)
                scroll_height = max(current_scroll[3], canvas_height) + extra_height
                
                # 스크롤 영역 업데이트
                self.thumb_canvas.config(scrollregion=(0, 0, scroll_width, scroll_height))
                
        except Exception as e:
            print(f"스크롤 영역 확인 중 오류: {e}")

    def handle_selection(self, event, index):
        # 드래그 시작 위치 저장
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        self.drag_start_index = index

        # 이미 여러 개가 선택된 상태에서 선택된 썸네일을 클릭(드래그 시작)하면 선택 상태 유지
        if len(self.selected_indices) > 1 and not (event.state & 0x0004) and not (event.state & 0x0001):
            if index in self.selected_indices:
                return

        if event.state & 0x0004:  # Ctrl
            if index in self.selected_indices:
                self.selected_indices.remove(index)
            else:
                self.selected_indices.add(index)
            self.last_clicked_index = index
        elif event.state & 0x0001:  # Shift
            if self.last_clicked_index is None:
                self.selected_indices = {index}
            else:
                start = min(self.last_clicked_index, index)
                end = max(self.last_clicked_index, index)
                self.selected_indices = set(range(start, end + 1))
        else:
            # Ctrl이나 Shift 없이 클릭하면 기존 선택 해제하고 새로 선택
            if len(self.selected_indices) > 1:
                self.selected_indices = {index}
            else:
                self.selected_indices = {index}
            self.last_clicked_index = index

        # 현재 페이지 인덱스 업데이트 및 미리보기 갱신
        self.current_page_index = index
        self.update_preview()
        self.update_selection_highlight()

    def on_drag(self, event, index):
        try:
            # 드래그 중인지 확인 (최소 이동 거리)
            if (abs(event.x - self.drag_data["x"]) > 5 or 
                abs(event.y - self.drag_data["y"]) > 5):
                
                # 선택된 페이지가 아니면 선택
                if index not in self.selected_indices:
                    self.selected_indices = {index}
                    self.update_selection_highlight()
                
                # 드래그 중임을 표시 (시각적 피드백 개선)
                for i in self.selected_indices:
                    if i < len(self.thumbnail_labels):
                        self.thumbnail_labels[i].config(relief="raised", bg="lightblue")
                
                # 드래그 중인 페이지들에 커서 변경
                for i in self.selected_indices:
                    if i < len(self.thumbnail_labels):
                        self.thumbnail_labels[i].config(cursor="fleur")
        except Exception as e:
            print(f"드래그 중 오류: {e}")

    def on_drop(self, event, index):
        try:
            if (abs(event.x - self.drag_data["x"]) > 5 or abs(event.y - self.drag_data["y"]) > 5):
                drop_target = self.get_simple_drop_target(event)
                if drop_target is not None:
                    # 선택된 페이지가 여러 개면, 드롭 위치가 선택된 영역 내부인지 체크
                    if self.selected_indices and drop_target in self.selected_indices:
                        return  # 자기 자신 위로 이동 방지
                    # drop_target이 선택된 영역보다 뒤에 있으면, 선택된 페이지 개수만큼 뒤로 보정
                    sorted_indices = sorted(self.selected_indices)
                    min_sel = sorted_indices[0]
                    max_sel = sorted_indices[-1]
                    if drop_target > max_sel:
                        drop_target = drop_target - len(self.selected_indices) + 1
                    # 선택된 페이지 전체 이동 (드래그 앤 드롭으로 페이지 이동 시 메시지 표시하지 않음)
                    self.move_pages_to_position(drop_target)
            for label in self.thumbnail_labels:
                label.config(relief="solid", bg="white", cursor="")
            self.drag_start_index = None
        except Exception as e:
            print(f"드롭 중 오류: {e}")
            for label in self.thumbnail_labels:
                label.config(relief="solid", bg="white", cursor="")
            self.drag_start_index = None
    
    def get_simple_drop_target(self, event):
        """간단한 드롭 타겟 찾기"""
        try:
            # 마우스 위치에 있는 위젯 찾기
            widget = event.widget.winfo_containing(event.x_root, event.y_root)
            
            # 썸네일 라벨인지 확인하고 인덱스 반환
            for i, label in enumerate(self.thumbnail_labels):
                if widget == label:
                    return i
            
            # 위젯을 찾지 못한 경우, 마우스 위치 기반으로 계산
            mouse_x = event.x_root
            mouse_y = event.y_root
            
            # 캔버스 내의 상대 위치 계산
            canvas_x = self.thumb_canvas.canvasx(mouse_x - self.thumb_canvas.winfo_rootx())
            canvas_y = self.thumb_canvas.canvasy(mouse_y - self.thumb_canvas.winfo_rooty())
            
            # 스크롤 위치 고려
            scroll_y = self.thumb_canvas.yview()[0] * self.thumb_scrollable_frame.winfo_height()
            adjusted_y = canvas_y + scroll_y
            
            if self.thumbnail_labels:
                # 그리드 레이아웃 계산
                sample_width = self.thumbnail_labels[0].winfo_reqwidth() + 10
                sample_height = self.thumbnail_labels[0].winfo_reqheight() + 10
                canvas_width = self.thumb_canvas.winfo_width()
                col_count = max(canvas_width // sample_width, 1)
                
                # 행과 열 계산
                row = int(adjusted_y // sample_height)
                col = int(canvas_x // sample_width)
                
                # 인덱스 계산
                index = row * col_count + col
                
                # 유효한 범위인지 확인
                if 0 <= index < len(self.thumbnail_labels):
                    return index
                elif index >= len(self.thumbnail_labels):
                    # 마지막 위치로 이동
                    return len(self.thumbnail_labels) - 1
            
            return None
        except Exception as e:
            print(f"간단한 드롭 타겟 찾기 오류: {e}")
            return None

    def get_drop_target(self, event):
        try:
            # 마우스 위치에 있는 위젯 찾기
            widget = event.widget.winfo_containing(event.x_root, event.y_root)
            
            # 썸네일 라벨인지 확인하고 인덱스 반환
            for i, label in enumerate(self.thumbnail_labels):
                if widget == label:
                    return i
            
            # 위젯을 찾지 못한 경우, 마우스 위치 기반으로 계산
            canvas_x = self.thumb_canvas.canvasx(event.x_root - self.thumb_canvas.winfo_rootx())
            canvas_y = self.thumb_canvas.canvasy(event.y_root - self.thumb_canvas.winfo_rooty())
            
            # 그리드 레이아웃을 기반으로 인덱스 계산
            if self.thumbnail_labels:
                sample_width = self.thumbnail_labels[0].winfo_reqwidth() + 10
                canvas_width = self.thumb_canvas.winfo_width()
                col_count = max(canvas_width // sample_width, 1)
                
                row = int(canvas_y // (self.thumbnail_labels[0].winfo_reqheight() + 10))
                col = int(canvas_x // sample_width)
                
                index = row * col_count + col
                if 0 <= index < len(self.thumbnail_labels):
                    return index
            
            return None
        except Exception as e:
            print(f"드롭 타겟 찾기 오류: {e}")
            return None

    def move_pages_to_position(self, target_index):
        if not self.selected_indices or target_index is None:
            return

        sorted_indices = sorted(self.selected_indices)
        n = len(self.doc)
        # 드롭 위치가 선택된 영역보다 뒤에 있으면, 보정
        if target_index > sorted_indices[-1]:
            target_index = target_index - len(sorted_indices) + 1

        # 현재 페이지 순서 리스트
        page_order = list(range(n))
        # 선택된 페이지 인덱스 제거
        for idx in reversed(sorted_indices):
            page_order.pop(idx)
        # 드롭 위치에 선택된 페이지 인덱스 삽입
        for i, idx in enumerate(sorted_indices):
            page_order.insert(target_index + i, idx)
        # 페이지 순서 재배치
        self.doc.select(page_order)
        # 선택 상태 업데이트
        self.selected_indices = set(range(target_index, target_index + len(sorted_indices)))
        
        # 현재 페이지 인덱스 조정
        if self.current_page_index in sorted_indices:
            # 선택된 페이지가 이동한 경우, 새로운 위치로 업데이트
            new_index = target_index + sorted_indices.index(self.current_page_index)
            self.current_page_index = new_index
        
        self.refresh_thumbnails()
        self.update_preview()
    
    def highlight_moved_pages(self, start_index, count):
        """이동된 페이지들을 잠시 하이라이트"""
        for i in range(start_index, start_index + count):
            if i < len(self.thumbnail_labels):
                self.thumbnail_labels[i].config(bg="lightgreen")
        
        # 1초 후 원래 색상으로 복원
        self.root.after(1000, self.reset_page_colors)
    
    def reset_page_colors(self):
        """모든 페이지 색상을 원래대로 복원"""
        for label in self.thumbnail_labels:
            label.config(bg="white")
    
    def on_enter(self, event, index):
        """마우스가 페이지 위에 올라왔을 때"""
        if index < len(self.thumbnail_labels):
            if index not in self.selected_indices:
                self.thumbnail_labels[index].config(bg="#f8f9fa")
    
    def on_leave(self, event, index):
        """마우스가 페이지에서 벗어났을 때"""
        if index < len(self.thumbnail_labels):
            if index not in self.selected_indices:
                self.thumbnail_labels[index].config(bg="white")
    
    def on_double_click(self, event, index):
        """더블클릭으로 페이지 이동"""
        if index in self.selected_indices:
            # 선택된 페이지들을 이동할 위치 입력 받기
            from tkinter import simpledialog
            target_pos = simpledialog.askinteger("페이지 이동", 
                                               f"선택된 {len(self.selected_indices)}개 페이지를 이동할 위치를 입력하세요 (1-{len(self.doc)}):",
                                               minvalue=1, maxvalue=len(self.doc))
            
            if target_pos is not None:
                target_index = target_pos - 1  # 0-based 인덱스로 변환
                self.move_pages_to_position(target_index)
        else:
            # 단일 페이지 이동
            from tkinter import simpledialog
            target_pos = simpledialog.askinteger("페이지 이동", 
                                               f"페이지 {index + 1}을 이동할 위치를 입력하세요 (1-{len(self.doc)}):",
                                               minvalue=1, maxvalue=len(self.doc))
            
            if target_pos is not None:
                target_index = target_pos - 1  # 0-based 인덱스로 변환
                self.selected_indices = {index}
                self.move_pages_to_position(target_index)

    def update_selection_highlight(self):
        for i, label in enumerate(self.thumbnail_labels):
            if i in self.selected_indices:
                label.config(highlightthickness=3, highlightbackground="#0078D4", bg="#e3f2fd")
            else:
                label.config(highlightthickness=0, bg="white")
        
        # 선택된 페이지 정보 업데이트
        self.update_selection_info()

    def update_selection_info(self):
        """선택된 페이지 정보 업데이트"""
        if not self.selected_indices:
            self.selection_info.config(text="선택된 페이지: 없음")
        else:
            if len(self.selected_indices) == 1:
                page_num = list(self.selected_indices)[0] + 1
                self.selection_info.config(text=f"선택된 페이지: {page_num}")
            else:
                sorted_indices = sorted(self.selected_indices)
                if len(sorted_indices) <= 5:
                    page_nums = [str(i + 1) for i in sorted_indices]
                    self.selection_info.config(text=f"선택된 페이지: {len(sorted_indices)}개 ({', '.join(page_nums)})")
                else:
                    page_nums = [str(sorted_indices[0] + 1), str(sorted_indices[1] + 1), 
                               "...", str(sorted_indices[-1] + 1)]
                    self.selection_info.config(text=f"선택된 페이지: {len(sorted_indices)}개 ({', '.join(page_nums)})")

    def delete_pages(self):
        if not self.selected_indices:
            return
            
        for idx in sorted(self.selected_indices, reverse=True):
            self.doc.delete_page(idx)
        
        max_idx = len(self.doc) - 1
        self.selected_indices = {min(i, max_idx) for i in self.selected_indices if i <= max_idx}
        
        # 현재 페이지 인덱스 조정
        if self.current_page_index >= len(self.doc):
            self.current_page_index = max(0, len(self.doc) - 1)
        
        self.refresh_thumbnails()
        self.update_preview()

    def rotate_pages(self, angle):
        """페이지 회전 (기존 함수 수정)"""
        if not self.selected_indices:
            messagebox.showwarning("경고", "회전할 페이지를 선택해주세요.")
            return
        
        for idx in self.selected_indices:
            current_rotation = self.doc[idx].rotation
            new_rotation = (current_rotation + angle) % 360
            self.doc[idx].set_rotation(new_rotation)
        
        self.refresh_thumbnails()
        self.update_preview()

    def move_pages(self, direction):
        if not self.selected_indices:
            return
            
        sorted_indices = sorted(self.selected_indices)
        updated = set()
        
        if direction < 0:
            for i in sorted_indices:
                if i > 0:
                    self.doc.move_page(i, i - 1)
                    updated.add(i - 1)
                else:
                    updated.add(i)
        else:
            for i in reversed(sorted_indices):
                if i < len(self.doc) - 1:
                    self.doc.move_page(i, i + 1)
                    updated.add(i + 1)
                else:
                    updated.add(i)
        
        self.selected_indices = updated
        self.refresh_thumbnails()
        self.update_preview()

    def save_pdf(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if path:
            self.doc.save(path)
            messagebox.showinfo("저장 완료", f"{path}로 저장됨")
    
    def move_selected_pages(self):
        """선택된 페이지들을 특정 위치로 이동"""
        if not self.selected_indices:
            messagebox.showwarning("경고", "이동할 페이지를 선택해주세요.")
            return
        
        # 이동할 위치 입력 받기
        from tkinter import simpledialog
        target_pos = simpledialog.askinteger("페이지 이동", 
                                           f"선택된 {len(self.selected_indices)}개 페이지를 이동할 위치를 입력하세요 (1-{len(self.doc)}):",
                                           minvalue=1, maxvalue=len(self.doc))
        
        if target_pos is not None:
            target_index = target_pos - 1  # 0-based 인덱스로 변환
            self.move_pages_to_position(target_index)

    def show_insert_blank_page_dialog(self):
        """빈페이지 삽입 다이얼로그 표시"""
        if not self.doc:
            messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
            return
        
        # 새 창 생성
        dialog = Toplevel(self.root)
        dialog.title("빈페이지 삽입")
        dialog.geometry("300x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 중앙 정렬
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        # 페이지 크기 선택
        tk.Label(dialog, text="페이지 크기를 선택하세요:", font=("맑은 고딕", 12)).pack(pady=10)
        
        # A4 가로/세로, A3 가로/세로 버튼들
        tk.Button(dialog, text="A4 가로 (210×297mm)", 
                 command=lambda: self.insert_blank_page("A4", "landscape", dialog)).pack(pady=5, fill=tk.X, padx=20)
        
        tk.Button(dialog, text="A4 세로 (210×297mm)", 
                 command=lambda: self.insert_blank_page("A4", "portrait", dialog)).pack(pady=5, fill=tk.X, padx=20)
        
        tk.Button(dialog, text="A3 가로 (297×420mm)", 
                 command=lambda: self.insert_blank_page("A3", "landscape", dialog)).pack(pady=5, fill=tk.X, padx=20)
        
        tk.Button(dialog, text="A3 세로 (297×420mm)", 
                 command=lambda: self.insert_blank_page("A3", "portrait", dialog)).pack(pady=5, fill=tk.X, padx=20)
        
        # 취소 버튼
        tk.Button(dialog, text="취소", command=dialog.destroy).pack(pady=10, fill=tk.X, padx=20)

    def insert_blank_page(self, page_size, orientation, dialog):
        """빈페이지 삽입"""
        try:
            # 페이지 크기 정의 (mm 단위)
            sizes = {
                "A4": {"width": 210, "height": 297},
                "A3": {"width": 297, "height": 420}
            }
            
            # 방향에 따른 크기 조정
            if orientation == "landscape":
                width = sizes[page_size]["height"]
                height = sizes[page_size]["width"]
            else:
                width = sizes[page_size]["width"]
                height = sizes[page_size]["height"]
            
            # mm를 포인트로 변환 (1mm = 2.83465 포인트)
            width_pt = width * 2.83465
            height_pt = height * 2.83465
            
            # 새 페이지 생성
            new_page = self.doc.new_page(width=width_pt, height=height_pt)
            
            # 현재 선택된 페이지 앞에 삽입
            if self.selected_indices:
                # 선택된 페이지 중 가장 앞쪽 위치에 삽입
                insert_pos = min(self.selected_indices)
                self.doc.move_page(len(self.doc) - 1, insert_pos)
            else:
                # 선택된 페이지가 없으면 맨 앞에 삽입
                self.doc.move_page(len(self.doc) - 1, 0)
            
            dialog.destroy()
            
            # 썸네일과 미리보기 새로고침
            self.refresh_thumbnails()
            self.update_preview()
            
        except Exception as e:
            messagebox.showerror("오류", f"빈페이지 삽입 중 오류가 발생했습니다: {str(e)}")

    def fit_page_to_screen(self):
        """페이지를 화면에 맞춤"""
        if not self.doc or self.current_page_index >= len(self.doc):
            return
        
        try:
            # 미리보기 패널의 크기 가져오기
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            if canvas_width <= 1 or canvas_height <= 1:
                return
            
            # 현재 페이지의 크기 가져오기
            page = self.doc[self.current_page_index]
            page_width = page.rect.width
            page_height = page.rect.height
            
            # 화면에 맞는 스케일 계산
            scale_x = (canvas_width - 20) / page_width  # 좌우 여백 10px씩
            scale_y = (canvas_height - 20) / page_height  # 상하 여백 10px씩
            
            # 더 작은 스케일을 사용하여 페이지가 화면을 벗어나지 않도록
            self.preview_scale = min(scale_x, scale_y)
            
            # 미리보기 업데이트
            self.update_preview()
            
        except Exception as e:
            messagebox.showerror("오류", f"페이지 맞춤 중 오류가 발생했습니다: {str(e)}")

    def delete_pages_with_key(self, event):
        """Delete 키로 페이지 삭제"""
        # 현재 포커스가 있는 위젯이 썸네일 캔버스인지 확인
        focused_widget = self.root.focus_get()
        
        # 썸네일 패널이나 미리보기 패널에 포커스가 있을 때만 삭제 실행
        if (focused_widget == self.thumb_canvas or 
            focused_widget == self.preview_canvas or
            focused_widget == self.root):
            
            if self.selected_indices:
                self.delete_pages()
            else:
                # 선택된 페이지가 없으면 현재 페이지 삭제
                if self.doc and self.current_page_index < len(self.doc):
                    self.selected_indices = {self.current_page_index}
                    self.delete_pages()

    def merge_pdf(self):
        """PDF 병합 기능 - 선택된 페이지 앞에 다른 PDF 파일 추가"""
        if not self.doc:
            messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
            return
        
        # 병합할 PDF 파일 선택
        merge_path = filedialog.askopenfilename(
            title="병합할 PDF 파일 선택",
            filetypes=[("PDF Files", "*.pdf")]
        )
        
        if not merge_path:
            return
        
        try:
            # 병합할 PDF 열기
            merge_doc = fitz.open(merge_path)
            
            if not merge_doc:
                messagebox.showerror("오류", "병합할 PDF 파일을 열 수 없습니다.")
                return
            
            # 삽입할 위치 결정
            if self.selected_indices:
                # 선택된 페이지 중 가장 앞쪽 위치에 삽입
                insert_pos = min(self.selected_indices)
            else:
                # 선택된 페이지가 없으면 맨 앞에 삽입
                insert_pos = 0
            
            # 병합할 PDF의 모든 페이지를 현재 문서에 추가
            added_pages = []
            for i in range(len(merge_doc)):
                try:
                    # 병합할 PDF의 페이지를 현재 문서에 복사
                    page = merge_doc[i]
                    
                    # 새 페이지 생성 (기존 페이지 크기 유지)
                    new_page = self.doc.new_page(width=page.rect.width, height=page.rect.height)
                    
                    # 페이지 내용 복사 (더 안전한 방법)
                    new_page.insert_image(new_page.rect, pixmap=page.get_pixmap())
                    
                    # 생성된 페이지를 임시로 저장
                    added_pages.append(len(self.doc) - 1)
                    
                except Exception as e:
                    print(f"페이지 {i} 복사 중 오류: {e}")
                    continue
            
            # 병합할 PDF 닫기
            merge_doc.close()
            
            if not added_pages:
                messagebox.showerror("오류", "병합할 페이지가 없습니다.")
                return
            
            # 추가된 페이지들을 원하는 위치로 이동
            for i, page_index in enumerate(added_pages):
                try:
                    self.doc.move_page(page_index, insert_pos + i)
                except Exception as e:
                    print(f"페이지 이동 중 오류: {e}")
                    continue
            
            # 선택 상태 업데이트 (새로 추가된 페이지들 선택)
            new_selection = set(range(insert_pos, insert_pos + len(added_pages)))
            self.selected_indices = new_selection
            
            # 썸네일과 미리보기 새로고침
            self.refresh_thumbnails()
            self.update_preview()
            
        except Exception as e:
            messagebox.showerror("오류", f"PDF 병합 중 오류가 발생했습니다: {str(e)}")
            print(f"PDF 병합 오류 상세: {e}")
        finally:
            # 병합 문서가 열려있다면 닫기
            try:
                if 'merge_doc' in locals() and merge_doc:
                    merge_doc.close()
            except:
                pass

    def clear_selection(self, event=None):
        """다중선택 해제"""
        if self.selected_indices:
            self.selected_indices.clear()
            self.update_selection_highlight()

    def select_all_pages(self, event=None):
        """모든 페이지 선택"""
        if self.doc:
            self.selected_indices = set(range(len(self.doc)))
            self.update_selection_highlight()

    def extract_text_directly(self):
        """바로 텍스트 추출 실행"""
        if not self.doc:
            messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
            return
        
        # 현재 선택된 페이지 또는 현재 페이지
        target_pages = list(self.selected_indices) if self.selected_indices else [self.current_page_index]
        
        if not target_pages:
            messagebox.showwarning("경고", "텍스트를 추출할 페이지를 선택해주세요.")
            return
        
        # 텍스트 추출 실행
        all_text = ""
        
        for i, page_idx in enumerate(target_pages):
            try:
                page_num = page_idx + 1
                page = self.doc[page_idx]
                
                # 기본 텍스트 추출
                text = page.get_text()
                if not text.strip():
                    text = "이 페이지에는 텍스트가 없습니다."
                
                all_text += f"=== 페이지 {page_num} ===\n{text}\n\n"
                
            except Exception as e:
                error_msg = f"페이지 {page_idx + 1} 처리 중 오류: {str(e)}"
                all_text += f"{error_msg}\n\n"
        
        # 결과를 새 창에 표시
        result_dialog = Toplevel(self.root)
        result_dialog.title(f"텍스트 추출 결과 - {len(target_pages)}개 페이지")
        result_dialog.geometry("800x600")
        result_dialog.transient(self.root)
        result_dialog.grab_set()
        
        # 중앙 정렬
        result_dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 100))
        
        # 제목
        title_label = tk.Label(result_dialog, text=f"{len(target_pages)}개 페이지에서 텍스트 추출 완료", 
                              font=("맑은 고딕", 14, "bold"))
        title_label.pack(pady=10)
        
        # 텍스트 표시 영역
        text_frame = tk.Frame(result_dialog)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("맑은 고딕", 10))
        text_scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=text_scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 텍스트 삽입
        text_widget.insert(tk.END, all_text)
        
        # 버튼 프레임
        button_frame = tk.Frame(result_dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # 복사 버튼
        copy_btn = ModernButton(button_frame, text="전체 텍스트 복사", 
                               command=lambda: copy_all_text(), bg="#3B82F6", fg="white")
        copy_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 저장 버튼
        save_btn = ModernButton(button_frame, text="텍스트 파일로 저장", 
                               command=lambda: save_text_file(), bg="#8B5CF6", fg="white")
        save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 닫기 버튼
        close_btn = ModernButton(button_frame, text="닫기", 
                                command=result_dialog.destroy, bg="#6B7280", fg="white")
        close_btn.pack(side=tk.RIGHT)
        
        def copy_all_text():
            """전체 텍스트를 클립보드에 복사"""
            try:
                result_dialog.clipboard_clear()
                result_dialog.clipboard_append(all_text)
                messagebox.showinfo("복사 완료", "텍스트가 클립보드에 복사되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"클립보드 복사 중 오류: {str(e)}")
        
        def save_text_file():
            """텍스트를 파일로 저장"""
            try:
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".txt",
                    filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
                )
                
                if file_path:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(all_text)
                    messagebox.showinfo("저장 완료", f"텍스트가 {file_path}에 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"파일 저장 중 오류: {str(e)}")

    def extract_text_from_page(self, page_index):
        """PDF 페이지에서 텍스트 추출"""
        try:
            if not self.doc or page_index >= len(self.doc):
                return ""
            
            page = self.doc[page_index]
            
            # 텍스트 추출 시도
            text = page.get_text()
            
            if text.strip():
                return text
            else:
                return "이 페이지에는 텍스트가 없습니다."
                
        except Exception as e:
            print(f"텍스트 추출 중 오류: {e}")
            return f"텍스트 추출 중 오류가 발생했습니다: {str(e)}"











    def show_text_extraction_dialog(self):
        """텍스트 추출 다이얼로그 표시"""
        if not self.doc:
            messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
            return
        
        # 현재 선택된 페이지 또는 현재 페이지
        target_pages = list(self.selected_indices) if self.selected_indices else [self.current_page_index]
        
        if not target_pages:
            messagebox.showwarning("경고", "텍스트를 추출할 페이지를 선택해주세요.")
            return
        
        # 새 창 생성
        dialog = Toplevel(self.root)
        dialog.title("텍스트 추출")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 중앙 정렬
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 100))
        
        # 메인 프레임
        main_frame = tk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 제목
        title_label = tk.Label(main_frame, text=f"선택된 {len(target_pages)}개 페이지에서 텍스트 추출", 
                              font=("맑은 고딕", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 옵션 프레임
        option_frame = tk.Frame(main_frame)
        option_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 기본 텍스트 추출 안내
        info_label = tk.Label(option_frame, text="PDF에서 텍스트를 추출합니다", 
                             font=("맑은 고딕", 10), fg="#6B7280")
        info_label.pack(side=tk.LEFT, padx=(0, 20))
        
        # 추출 버튼
        extract_btn = ModernButton(option_frame, text="텍스트 추출", 
                                  command=lambda: extract_text(), bg="#10B981", fg="white")
        extract_btn.pack(side=tk.RIGHT)
        
        # 텍스트 표시 영역
        text_frame = tk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # 텍스트 위젯과 스크롤바
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("맑은 고딕", 10))
        text_scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=text_scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 하단 버튼 프레임
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 복사 버튼
        copy_btn = ModernButton(button_frame, text="전체 텍스트 복사", 
                               command=lambda: copy_all_text(), bg="#3B82F6", fg="white")
        copy_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 저장 버튼
        save_btn = ModernButton(button_frame, text="텍스트 파일로 저장", 
                               command=lambda: save_text_file(), bg="#8B5CF6", fg="white")
        save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 닫기 버튼
        close_btn = ModernButton(button_frame, text="닫기", 
                                command=dialog.destroy, bg="#6B7280", fg="white")
        close_btn.pack(side=tk.RIGHT)
        
        def extract_text():
            """텍스트 추출 실행"""
            text_widget.delete(1.0, tk.END)
            text_widget.insert(tk.END, "텍스트 추출 중...\n\n")
            dialog.update()
            
            all_text = ""
            
            for i, page_idx in enumerate(target_pages):
                try:
                    page_num = page_idx + 1
                    text_widget.insert(tk.END, f"=== 페이지 {page_num} ===\n")
                    
                    # 기본 텍스트 추출
                    text = self.doc[page_idx].get_text()
                    if not text.strip():
                        text = "이 페이지에는 텍스트가 없습니다."
                    
                    text_widget.insert(tk.END, f"{text}\n\n")
                    all_text += f"=== 페이지 {page_num} ===\n{text}\n\n"
                    
                    dialog.update()
                    
                except Exception as e:
                    error_msg = f"페이지 {page_idx + 1} 처리 중 오류: {str(e)}"
                    text_widget.insert(tk.END, f"{error_msg}\n\n")
                    all_text += f"{error_msg}\n\n"
            
            text_widget.insert(tk.END, "텍스트 추출이 완료되었습니다.")
            
            # 전역 변수로 저장 (복사/저장용)
            dialog.extracted_text = all_text
        
        def copy_all_text():
            """전체 텍스트를 클립보드에 복사"""
            try:
                if hasattr(dialog, 'extracted_text') and dialog.extracted_text:
                    dialog.clipboard_clear()
                    dialog.clipboard_append(dialog.extracted_text)
                    messagebox.showinfo("복사 완료", "텍스트가 클립보드에 복사되었습니다.")
                else:
                    messagebox.showwarning("경고", "먼저 텍스트를 추출해주세요.")
            except Exception as e:
                messagebox.showerror("오류", f"클립보드 복사 중 오류: {str(e)}")
        
        def save_text_file():
            """텍스트를 파일로 저장"""
            try:
                if hasattr(dialog, 'extracted_text') and dialog.extracted_text:
                    file_path = filedialog.asksaveasfilename(
                        defaultextension=".txt",
                        filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
                    )
                    
                    if file_path:
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(dialog.extracted_text)
                        messagebox.showinfo("저장 완료", f"텍스트가 {file_path}에 저장되었습니다.")
                else:
                    messagebox.showwarning("경고", "먼저 텍스트를 추출해주세요.")
            except Exception as e:
                messagebox.showerror("오류", f"파일 저장 중 오류: {str(e)}")

    def show_quick_text_extraction(self):
        """빠른 텍스트 추출 (현재 페이지만)"""
        if not self.doc:
            messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
            return
        
        try:
            # 현재 페이지에서 텍스트 추출
            text = self.doc[self.current_page_index].get_text()
            
            if not text.strip():
                text = "이 페이지에서 텍스트를 추출할 수 없습니다."
            
            # 결과를 새 창에 표시
            result_dialog = Toplevel(self.root)
            result_dialog.title(f"페이지 {self.current_page_index + 1} 텍스트")
            result_dialog.geometry("600x400")
            result_dialog.transient(self.root)
            result_dialog.grab_set()
            
            # 중앙 정렬
            result_dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 150, self.root.winfo_rooty() + 150))
            
            # 제목
            title_label = tk.Label(result_dialog, text=f"페이지 {self.current_page_index + 1} 텍스트", 
                                  font=("맑은 고딕", 12, "bold"))
            title_label.pack(pady=10)
            
            # 텍스트 표시 영역
            text_frame = tk.Frame(result_dialog)
            text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
            
            text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("맑은 고딕", 10))
            text_scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
            text_widget.configure(yscrollcommand=text_scrollbar.set)
            
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 텍스트 삽입
            text_widget.insert(tk.END, text)
            
            # 버튼 프레임
            button_frame = tk.Frame(result_dialog)
            button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
            
            # 복사 버튼
            copy_btn = ModernButton(button_frame, text="텍스트 복사", 
                                   command=lambda: copy_text(), bg="#3B82F6", fg="white")
            copy_btn.pack(side=tk.LEFT, padx=(0, 10))
            
            # 닫기 버튼
            close_btn = ModernButton(button_frame, text="닫기", 
                                    command=result_dialog.destroy, bg="#6B7280", fg="white")
            close_btn.pack(side=tk.RIGHT)
            
            def copy_text():
                """텍스트를 클립보드에 복사"""
                try:
                    result_dialog.clipboard_clear()
                    result_dialog.clipboard_append(text)
                    messagebox.showinfo("복사 완료", "텍스트가 클립보드에 복사되었습니다.")
                except Exception as e:
                    messagebox.showerror("오류", f"클립보드 복사 중 오류: {str(e)}")
                    
        except Exception as e:
            messagebox.showerror("오류", f"텍스트 추출 중 오류가 발생했습니다: {str(e)}")

if __name__ == "__main__":
    if DRAG_DROP_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    root.geometry("1400x900")  # 더 넓은 창 크기로 설정
    app = PDFEditorApp(root)
    root.mainloop()