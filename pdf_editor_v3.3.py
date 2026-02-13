import fitz
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, ttk
from PIL import Image, ImageTk
from functools import partial
import os
import tempfile
import sys
import uuid
import subprocess
import re
import base64

import json
import time
import gc

# Kunhwa PDF Editor v3.3 - Undo/Redo + 진행률 + GoToPage + 상태표시줄 + 최근파일 + 모던UI
VERSION = "v3.3"

# 사용자 정보는 외부 암호화 파일(users.json.enc)에서 관리

# tkinterdnd2 라이브러리 임포트 (윈도우 드래그 앤 드롭 지원)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    try:
        from tkinterdnd2 import DND_TEXT
    except Exception:
        DND_TEXT = 'text/plain'
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False
    print("tkinterdnd2 라이브러리가 설치되지 않았습니다. 드래그 앤 드롭 기능을 사용할 수 없습니다.")
    print("설치 방법: pip install tkinterdnd2")

# 멀티 창/교차 드래그를 위한 전역 레지스트리
OPEN_EDITORS = []

# ─────────────────────────────────────────────
# 앱 데이터 디렉토리 헬퍼
# ─────────────────────────────────────────────
def _get_app_data_dir():
    """앱 설정 파일 저장 경로 반환 (APPDATA 기반)"""
    try:
        base = os.environ.get('APPDATA', os.path.expanduser('~'))
        d = os.path.join(base, 'KunhwaPDFEditor')
        os.makedirs(d, exist_ok=True)
        return d
    except Exception:
        return os.getcwd()

# ─────────────────────────────────────────────
# Undo / Redo 매니저
# ─────────────────────────────────────────────
class UndoManager:
    """PDF 편집 작업의 Undo/Redo를 관리 (최대 10단계)"""
    def __init__(self, max_history=10):
        self._undo_stack = []   # [(action_name, pdf_bytes), ...]
        self._redo_stack = []
        self.max_history = max_history

    def save_state(self, doc, action_name=""):
        """현재 PDF 상태를 undo 스택에 저장"""
        if doc is None:
            return
        try:
            pdf_bytes = doc.tobytes(deflate=True)
            self._undo_stack.append((action_name, pdf_bytes))
            # 스택 크기 제한
            while len(self._undo_stack) > self.max_history:
                self._undo_stack.pop(0)
            # 새 작업을 하면 redo 스택은 초기화
            self._redo_stack.clear()
        except Exception as e:
            print(f"Undo 상태 저장 실패: {e}")

    def undo(self, doc):
        """이전 상태로 되돌리기. 복원할 PDF bytes 반환, 없으면 None"""
        if not self._undo_stack:
            return None
        try:
            # 현재 상태를 redo 스택에 저장
            if doc is not None:
                current_bytes = doc.tobytes(deflate=True)
                self._redo_stack.append(("redo", current_bytes))
            action_name, prev_bytes = self._undo_stack.pop()
            print(f"Undo: '{action_name}' 작업 되돌리기")
            return prev_bytes
        except Exception as e:
            print(f"Undo 실패: {e}")
            return None

    def redo(self, doc):
        """되돌린 작업 다시 실행. 복원할 PDF bytes 반환, 없으면 None"""
        if not self._redo_stack:
            return None
        try:
            # 현재 상태를 undo 스택에 저장
            if doc is not None:
                current_bytes = doc.tobytes(deflate=True)
                self._undo_stack.append(("undo", current_bytes))
            _, redo_bytes = self._redo_stack.pop()
            print("Redo: 작업 다시 실행")
            return redo_bytes
        except Exception as e:
            print(f"Redo 실패: {e}")
            return None

    @property
    def can_undo(self):
        return len(self._undo_stack) > 0

    @property
    def can_redo(self):
        return len(self._redo_stack) > 0

    def clear(self):
        self._undo_stack.clear()
        self._redo_stack.clear()

# ─────────────────────────────────────────────
# 최근 파일 매니저
# ─────────────────────────────────────────────
class RecentFilesManager:
    """최근 열었던 파일 목록 관리 (최대 5개)"""
    MAX_FILES = 5

    def __init__(self):
        self._path = os.path.join(_get_app_data_dir(), 'recent_files.json')
        self._files = self._load()

    def _load(self):
        try:
            if os.path.exists(self._path):
                with open(self._path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return [p for p in data if os.path.exists(p)][:self.MAX_FILES]
        except Exception:
            pass
        return []

    def _save(self):
        try:
            with open(self._path, 'w', encoding='utf-8') as f:
                json.dump(self._files, f, ensure_ascii=False)
        except Exception as e:
            print(f"최근 파일 저장 실패: {e}")

    def add(self, file_path):
        """파일 경로 추가 (중복이면 맨 앞으로 이동)"""
        abs_path = os.path.abspath(file_path)
        if abs_path in self._files:
            self._files.remove(abs_path)
        self._files.insert(0, abs_path)
        self._files = self._files[:self.MAX_FILES]
        self._save()

    @property
    def files(self):
        return list(self._files)

    def remove(self, file_path):
        """파일 경로 제거"""
        abs_path = os.path.abspath(file_path)
        if abs_path in self._files:
            self._files.remove(abs_path)
            self._save()

# ─────────────────────────────────────────────
# 프로그레스 바 (v3.3 신규)
# ─────────────────────────────────────────────
class ProgressIndicator:
    """대화상자 형태의 프로그레스 바"""
    def __init__(self, parent, title="처리 중", maximum=100):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("350x150")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        
        # 화면 중앙 배치
        self.top.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 175
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - 75
        self.top.geometry(f"+{x}+{y}")
        
        self.top.configure(bg="white")
        container = tk.Frame(self.top, bg="white", padx=20, pady=20)
        container.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(container, text=title, font=("맑은 고딕", 11, "bold"), bg="white", fg="#2563EB").pack(pady=(0, 10))
        
        self.progress = ttk.Progressbar(container, orient=tk.HORIZONTAL, length=300, mode='determinate', maximum=maximum)
        self.progress.pack(pady=10)
        
        self.label = tk.Label(container, text="준비 중...", font=("맑은 고딕", 9), bg="white", fg="#4B5563")
        self.label.pack()
        
    def update(self, value, text=None):
        self.progress['value'] = value
        if text:
            self.label.config(text=text)
        else:
            percent = int((value / self.progress['maximum']) * 100)
            self.label.config(text=f"진행률: {percent}%")
        self.top.update()
        
    def close(self):
        try:
            self.top.grab_release()
            self.top.destroy()
        except:
            pass


def _build_launch_command(extra_args=None):
    """현재 실행 환경(스크립트/EXE)에 맞춘 새 프로세스 실행 명령 생성"""
    extra_args = extra_args or []
    if getattr(sys, "frozen", False):
        base_cmd = [sys.executable]
    else:
        base_cmd = [sys.executable, os.path.abspath(__file__)]
    return base_cmd + extra_args

def get_mac_address():
    """시스템의 맥어드레스를 가져오는 함수 (최적화)"""
    try:
        # Windows에서 맥어드레스 가져오기 (캐싱 적용)
        if hasattr(get_mac_address, '_cached_mac'):
            return get_mac_address._cached_mac
            
        result = subprocess.run(['ipconfig', '/all'], capture_output=True, text=True, encoding='cp949')
        if result.returncode == 0:
            # 물리적 주소(Physical Address) 찾기
            lines = result.stdout.split('\n')
            for line in lines:
                if '물리적 주소' in line or 'Physical Address' in line:
                    mac_match = re.search(r'([0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2})', line, re.IGNORECASE)
                    if mac_match:
                        mac = mac_match.group(1).upper()
                        get_mac_address._cached_mac = mac  # 캐싱
                        return mac
        
        # 대안 방법: uuid 모듈 사용
        mac = uuid.getnode()
        mac_address = ':'.join(['{:02x}'.format((mac >> elements) & 0xff) for elements in range(0,2*6,2)][::-1])
        # 콜론을 하이픈으로 변환
        mac = mac_address.replace(':', '-').upper()
        get_mac_address._cached_mac = mac  # 캐싱
        return mac
        
    except Exception as e:
        print(f"맥어드레스 가져오기 실패: {e}")
        return None

def encrypt_users_data(json_content, password="KUNHWA2025"):
    """JSON 내용을 암호화"""
    try:
        # JSON을 바이트로 변환
        json_bytes = json_content.encode('utf-8')
        
        # 간단한 XOR 암호화 + Base64 인코딩
        encrypted = bytearray()
        for i, byte in enumerate(json_bytes):
            key_byte = ord(password[i % len(password)])
            encrypted.append(byte ^ key_byte)
        
        # Base64로 인코딩하여 텍스트로 변환
        return base64.b64encode(encrypted).decode('utf-8')
    except Exception as e:
        print(f"암호화 실패: {e}")
        return None

def decrypt_users_data(encrypted_text, password="KUNHWA2025"):
    """암호화된 텍스트를 복호화"""
    try:
        # Base64 디코딩
        encrypted = base64.b64decode(encrypted_text.encode('utf-8'))
        
        # XOR 복호화
        decrypted = bytearray()
        for i, byte in enumerate(encrypted):
            key_byte = ord(password[i % len(password)])
            decrypted.append(byte ^ key_byte)
        
        return decrypted.decode('utf-8')
    except Exception as e:
        print(f"복호화 실패: {e}")
        return None

def save_encrypted_users(users_data, filename="users.json.enc"):
    """사용자 정보를 암호화하여 파일로 저장"""
    try:
        # 실행 디렉토리 기준 경로 사용(Exe 패킹 후에도 동일)
        def _get_app_dir():
            try:
                if getattr(sys, 'frozen', False):
                    return os.path.dirname(sys.executable)
                return os.path.dirname(os.path.abspath(__file__))
            except Exception:
                return os.getcwd()

        if not filename or filename == "users.json.enc":
            filename = os.path.join(_get_app_dir(), "users.json.enc")

        encrypted = encrypt_users_data(users_data)
        if encrypted:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(encrypted)
            print(f"암호화된 사용자 정보 저장 완료: {filename}")
            return True
        return False
    except Exception as e:
        print(f"사용자 정보 저장 실패: {e}")
        return False

def load_encrypted_users(filename="users.json.enc"):
    """암호화된 파일에서 사용자 정보 로드 (EXE 설치 디렉토리 기준)"""
    try:
        # 실행 디렉토리 기준 경로 사용(Exe 패킹 후에도 동일)
        def _get_app_dir():
            try:
                if getattr(sys, 'frozen', False):
                    return os.path.dirname(sys.executable)
                return os.path.dirname(os.path.abspath(__file__))
            except Exception:
                return os.getcwd()

        if not filename or filename == "users.json.enc":
            filename = os.path.join(_get_app_dir(), "users.json.enc")

        if not os.path.exists(filename):
            print(f"사용자 정보 파일이 없습니다: {filename}")
            return None
        
        with open(filename, 'r', encoding='utf-8') as f:
            encrypted = f.read()
        
        decrypted = decrypt_users_data(encrypted)
        if decrypted:
            return json.loads(decrypted)
        return None
    except Exception as e:
        print(f"사용자 정보 로드 실패: {e}")
        return None

def create_default_users_file():
    """기본 사용자 정보 파일 생성"""
    default_users = {
        "users": [
            {"mac": "80-E8-2C-EF-97-E0", "name": "장태웅", "role": "admin"},
            {"mac": "6C-0B-5E-42-ED-70", "name": "최건호", "role": "user"},
            {"mac": "80-E8-2C-EF-9E-4E", "name": "문석환", "role": "user"},
            {"mac": "E0-73-E7-BB-62-DE", "name": "허재혁", "role": "user"},
            {"mac": "6C-0B-5E-42-EC-3D", "name": "유청산", "role": "user"},
            {"mac": "BC-E9-2F-A1-AD-6C", "name": "김지환", "role": "user"},
            {"mac": "2C-58-B9-92-F5-CB", "name": "김대영", "role": "user"},
            {"mac": "40-1A-58-F7-76-FF", "name": "노트북(HP)", "role": "user"},
            {"mac": "6C-0B-5E-42-EB-FE", "name": "왕세환", "role": "user"}
        ],
        "last_updated": "2025-01-27",
        "version": "1.0"
    }
    
    # JSON 형식으로 변환
    json_content = json.dumps(default_users, ensure_ascii=False, indent=2)
    
    # 암호화하여 저장
    if save_encrypted_users(json_content):
        print("기본 사용자 정보 파일 생성 완료")
        return True
    return False

def check_authorization():
    """사용자 인증 확인 - 암호화된 파일에서 로드"""
    current_mac = get_mac_address()
    if not current_mac:
        messagebox.showerror("인증 오류", "시스템 맥어드레스를 확인할 수 없습니다.")
        return False
    
    print(f"현재 시스템 맥어드레스: {current_mac}")
    
    # 설치 디렉토리에서 인증 파일 로드
    # (관리자가 설치 폴더에 users.json.enc를 직접 배치해야 함)
    users_data = load_encrypted_users()
    
    if not users_data:
        # 인증 파일이 없는 경우 - 설치 경로 안내
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
        
        print(f"사용자 정보 파일이 없습니다. 필요 위치: {app_dir}")
        messagebox.showerror("인증 오류", 
                            "사용자 정보 파일(users.json.enc)을 찾을 수 없습니다.\n\n"
                            f"파일 위치: {app_dir}\n\n"
                            "관리자에게 인증 파일을 받아\n"
                            "위 폴더에 넣어주세요.")
        return False
    
    # 사용자 검증
    if "users" in users_data:
        for user in users_data["users"]:
            if user["mac"] == current_mac:
                print(f"인증 성공: {user['name']} ({user['role']})")
                return True
    
    # 인증 실패 시 오류 메시지
    messagebox.showerror("접근 거부", 
                        "이 프로그램은 허용된 사용자만 실행할 수 있습니다.\n\n"
                        "관리자에게 문의하세요.")
    return False

# 기본 텍스트 추출 기능만 사용

class ModernButton(tk.Button):
    """모던한 디자인의 버튼 클래스"""
    def __init__(self, parent, **kwargs):
        # 기본 스타일 설정
        default_style = {
            'font': ('맑은 고딕', 8, 'bold'),
            'relief': 'raised',
            'borderwidth': 1,
            'padx': 6,
            'pady': 3,
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
        # 이 창 고유 ID (교차 드래그 식별자)
        self.window_id = str(id(self))
        # 전역 레지스트리에 자신 등록 (교차 드래그 타겟 검색 용도)
        try:
            OPEN_EDITORS.append(self)
        except Exception:
            pass
        # 창 종료 시 레지스트리 정리
        try:
            self.root.protocol("WM_DELETE_WINDOW", self._on_close_window)
        except Exception:
            pass
        
        # ── 사용자 데이터 캐싱 (1회만 로드) ──
        current_mac = get_mac_address()
        self._cached_users_data = load_encrypted_users()
        user_name = ""
        self._current_user_role = "user"
        if self._cached_users_data and "users" in self._cached_users_data:
            for user in self._cached_users_data["users"]:
                if user["mac"] == current_mac:
                    user_name = user["name"]
                    self._current_user_role = user.get("role", "user")
                    break
        self._cached_user_name = user_name
        
        # 제목에 사용자 정보와 버전 포함
        if user_name:
            self.root.title(f"Kunhwa PDF Editor {VERSION} - {user_name}")
        else:
            self.root.title(f"Kunhwa PDF Editor {VERSION}")
        
        # 인증 정보 터미널 출력
        print(f"=== Kunhwa PDF Editor {VERSION} ===")
        print(f"인증된 사용자: {user_name}")
        print("프로그램이 성공적으로 시작되었습니다.")
        print("=" * 35)
        
        # 윈도우 스타일 설정
        self.root.configure(bg='#f8f9fa')
        
        # 상태 값들 (최적화)
        self.doc = None
        self.thumb_scale = 0.20  # 기본값 0.20
        self.preview_scale = 1.00  # 기본값 1.00
        self.current_page_index = 0  # 미리보기 표시할 페이지 인덱스
        self.selected_indices = set()
        self._zoom_target = 'thumbs'  # 줌 대상 패널
        
        # 성능 설정 관련 (새로 추가)
        self.performance_mode = "balanced"  # 기본값: 균형 모드
        
        # 썸네일 관련 (최적화)
        self.thumbnails = []
        self.thumbnail_labels = []
        self.thumbnail_frames = []
        self.last_clicked_index = None
        self.drag_start_index = None
        self.drag_data = {"x": 0, "y": 0, "item": None}
        
        # 성능 최적화를 위한 고급 캐싱 시스템
        self._thumbnail_cache = {}
        self._preview_cache = {}
        self._last_update_time = 0
        self._cache_size_limit = 100  # 캐시 크기 제한
        self._cache_hits = 0
        self._cache_misses = 0
        # 페이지 클립보드 (교차 창 공유)
        self.page_clipboard = []
        self.page_clipboard_bytes = None
        
        # ── v3.3 신규: Undo/Redo 매니저 ──
        self.undo_manager = UndoManager(max_history=10)
        
        # ── v3.3 신규: 최근 파일 매니저 ──
        self.recent_files_manager = RecentFilesManager()
        
        self.setup_ui()
        self.bind_events()

    def setup_drag_drop(self):
        """드래그 앤 드롭 설정"""
        if not DRAG_DROP_AVAILABLE:
            return
        
        try:
            # 공통 드롭 라우터 등록 (파일/텍스트 모두 지원)
            def register_targets(widget):
                try:
                    widget.drop_target_register(DND_FILES, DND_TEXT)
                except Exception:
                    try:
                        widget.drop_target_register(DND_FILES)
                        widget.drop_target_register(DND_TEXT)
                    except Exception:
                        pass
                try:
                    widget.dnd_bind('<<Drop>>', self.on_generic_drop)
                except Exception:
                    pass

            # 메인 윈도우 / 썸네일 캔버스 / 미리보기 캔버스 모두 등록
            register_targets(self.root)
            
            register_targets(self.thumb_canvas)
            register_targets(self.preview_canvas)
            
            print("드래그 앤 드롭 기능이 활성화되었습니다.")
        except Exception as e:
            print(f"드래그 앤 드롭 설정 중 오류: {e}")

    def setup_menu_bar(self):
        """메뉴바 설정 (새로 추가)"""
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        
        # 파일 메뉴
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="새 창 (빈)", accelerator="Ctrl+N", command=self.new_window)
        file_menu.add_command(label="파일로 새 창...", accelerator="Ctrl+Shift+N", command=self.new_window_with_file)
        file_menu.add_separator()
        file_menu.add_command(label="PDF 열기", command=self.open_pdf, accelerator="Ctrl+O")
        file_menu.add_command(label="PDF 저장", command=self.save_pdf, accelerator="Ctrl+S")
        file_menu.add_command(label="선택 페이지 저장", command=self.save_selected_pages)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.root.quit, accelerator="Alt+F4")
        
        # 편집 메뉴
        self.edit_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="편집", menu=self.edit_menu)
        self.edit_menu.add_command(label="되돌리기 (Undo)", accelerator="Ctrl+Z", command=self.perform_undo)
        self.edit_menu.add_command(label="다시실행 (Redo)", accelerator="Ctrl+Y", command=self.perform_redo)
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="페이지 이동...", accelerator="Ctrl+G", command=self.show_goto_page_dialog)
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="페이지 회전 (우로90°)", accelerator="Ctrl+]", command=lambda: self.rotate_pages(90))
        self.edit_menu.add_command(label="페이지 회전 (좌로90°)", accelerator="Ctrl+[", command=lambda: self.rotate_pages(-90))
        self.edit_menu.add_command(label="빈페이지 삽입", command=self.show_insert_blank_page_dialog)
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="텍스트 추출", command=self.extract_text_directly)
        
        # 최근 파일 서브메뉴
        self._recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.insert_cascade(5, label="최근 파일", menu=self._recent_menu)
        self._refresh_recent_files_menu()
        
        # 성능 설정 메뉴 (새로 추가)
        performance_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="성능 설정", menu=performance_menu)
        
        # 성능 모드 서브메뉴
        performance_menu.add_command(label="🚀 고성능 모드 (권장)", 
                                   command=lambda: self.set_performance_mode("high"))
        performance_menu.add_command(label="⚖️ 균형 모드", 
                                   command=lambda: self.set_performance_mode("balanced"))
        performance_menu.add_command(label="🎨 고품질 모드", 
                                   command=lambda: self.set_performance_mode("quality"))
        
        performance_menu.add_separator()
        
        # 현재 설정 표시
        self.current_mode_label = tk.StringVar()
        self.current_mode_label.set("현재: ⚖️ 균형 모드")
        performance_menu.add_command(label="현재: ⚖️ 균형 모드", state="disabled")
        
        # 사용자 관리 메뉴 (관리자만 표시)
        self.users_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="사용자 관리", menu=self.users_menu)
        self.users_menu.add_command(label="사용자 목록 보기", command=self.show_users_list)
        self.users_menu.add_command(label="사용자 추가", command=self.add_user)
        self.users_menu.add_command(label="사용자 제거", command=self.remove_user)
        self.users_menu.add_separator()
        self.users_menu.add_command(label="JSON 파일 편집", command=self.edit_users_json)
        self.users_menu.add_command(label="사용자 백업", command=self.backup_users)
        self.users_menu.add_command(label="백업 복원", command=self.restore_users_backup)
        
        # 사용자 권한에 따라 메뉴 활성화/비활성화
        self._update_user_menu_visibility()
        
        # 도움말 메뉴
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="도움말", menu=help_menu)
        help_menu.add_command(label="사용법", command=self.show_help)
        help_menu.add_command(label="정보", command=self.show_about)

        # (요청) 창 메뉴 제거

    def _set_all_user_menu_state(self, state):
        """사용자 관리 메뉴의 모든 항목 상태 설정 (v3.3)"""
        items = ["사용자 목록 보기", "사용자 추가", "사용자 제거", "JSON 파일 편집", "사용자 백업", "백업 복원"]
        for item in items:
            try:
                self.users_menu.entryconfig(item, state=state)
            except:
                pass

    def _update_user_menu_visibility(self):
        """사용자 권한에 따라 메뉴 가시성 업데이트 (v3.3 최적화: 캐싱 활용)"""
        try:
            # __init__에서 캐싱된 역할 정보 활용
            if hasattr(self, '_current_user_role') and self._current_user_role == "admin":
                self._set_all_user_menu_state("normal")
            else:
                self._set_all_user_menu_state("disabled")
        except Exception as e:
            print(f"메뉴 가시성 업데이트 실패: {e}")
            try:
                self._set_all_user_menu_state("disabled")
            except:
                pass

    # (중복 정의 제거) open_pdf_from_path, create_pdf_from_image는 아래 최신 구현을 사용


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
                first_file = supported_files[0]
                file_ext = os.path.splitext(first_file)[1].lower()
                
                try:
                    if file_ext == '.pdf':
                        success = self.open_pdf_from_path(first_file)
                        if not success:
                            messagebox.showerror("오류", f"PDF 파일을 열 수 없습니다:\n{os.path.basename(first_file)}")
                            return
                    else:
                        # 첫 번째 파일이 이미지인 경우 새 PDF 생성
                        self.create_pdf_from_image(first_file)
                    
                    # 새로 열린 PDF의 경우 드롭 위치를 마지막으로 설정
                    drop_target = len(self.doc) if self.doc else 0
                    print(f"새 PDF 열기 후 마지막 위치로 설정: {drop_target}")
                    
                    # 첫 번째 파일은 이미 처리했으므로 제외
                    supported_files = supported_files[1:]
                    
                except Exception as e:
                    error_msg = f"파일 열기 실패:\n{os.path.basename(first_file)}\n{str(e)}"
                    messagebox.showerror("파일 열기 오류", error_msg)
                    print(f"드롭 파일 열기 실패: {e}")
                    return
            
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

    def on_generic_drop(self, event):
        """파일/텍스트 드롭 라우팅"""
        try:
            data = getattr(event, 'data', '') or ''
            # 텍스트 페이로드(PDFTHUMB::) → 창 간 이동 처리 (window_id::index 포함)
            if isinstance(data, str) and 'PDFTHUMB::' in data:
                # 드롭 위치 계산 (썸네일 위면 그 위치, 아니면 마지막)
                drop_pos = -1
                try:
                    if event.widget == self.thumb_canvas:
                        drop_pos = self.get_thumbnail_drop_position(event)
                except Exception:
                    pass
                if drop_pos >= 0:
                    self.drop_target_index = drop_pos
                return self.on_interwindow_drop(event)
            # 파일 드롭로 처리 (PDF, 이미지)
            return self.on_drop_file(event)
        except Exception as e:
            print(f"드롭 라우팅 오류: {e}")

    def open_pdf_from_path(self, file_path):
        """경로로부터 PDF 열기 - 개선된 버전"""
        try:
            # 파일 존재 여부 확인
            if not os.path.exists(file_path):
                print(f"파일이 존재하지 않습니다: {file_path}")
                return False
            
            # 파일 크기 확인
            if os.path.getsize(file_path) == 0:
                print(f"파일이 비어있습니다: {file_path}")
                return False
            
            # 기존 문서가 있다면 정리
            if self.doc:
                try:
                    self.doc.close()
                except:
                    pass
                self.doc = None
            
            # 새 PDF 문서 열기
            self.doc = fitz.open(file_path)
            
            # 문서 유효성 검사
            if not self.doc or len(self.doc) == 0:
                print(f"PDF 파일을 읽을 수 없거나 페이지가 없습니다: {file_path}")
                if self.doc:
                    self.doc.close()
                    self.doc = None
                return False
            
            # 상태 초기화
            self.current_page_index = 0
            self.selected_indices.clear()
            
            # UI 업데이트
            self.refresh_thumbnails()
            self.update_preview()
            
            # 파일명 표시
            try:
                filename = os.path.basename(file_path)
                if hasattr(self, 'thumb_filename_label'):
                    self.thumb_filename_label.config(text=filename)
                if hasattr(self, 'root') and self.root:
                    self.root.title(f"Kunhwa PDF Editor {VERSION} - {filename}")
            except Exception as e:
                print(f"파일명 표시 중 오류: {e}")
            
            print(f"PDF 열기 성공: {file_path} ({len(self.doc)}페이지)")
            
            # v3.3: 최근 파일 목록에 추가 & 상태표시줄 업데이트
            self.recent_files_manager.add(file_path)
            self._refresh_recent_files_menu()
            self._update_status_bar()
            
            return True
            
        except Exception as e:
            print(f"PDF 열기 실패: {e}")
            # 오류 발생 시 문서 상태 정리
            if self.doc:
                try:
                    self.doc.close()
                except:
                    pass
                self.doc = None
            return False

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
                print("메인 윈도우에 드롭됨: 마지막 위치로 설정")
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
                    print("드롭 위치 계산: 기본 위치를 마지막으로 설정")
                    return -1  # 기본 위치도 마지막으로 설정
            
            # 썸네일이 없는 경우도 마지막 위치로 처리
            print("썸네일이 없음: 마지막 위치로 설정")
            return -1
        except Exception as e:
            print(f"썸네일 드롭 위치 계산 오류: {e}")
            return -1  # 오류 발생 시에도 마지막 위치로 설정

    def merge_pdf_from_path_with_position(self, file_path, insert_pos):
        """지정된 위치에 PDF 병합"""
        if self.doc:
            self.undo_manager.save_state(self.doc)
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
        if self.doc:
            self.undo_manager.save_state(self.doc)
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
        # 메뉴바 추가 (최상단)
        self.setup_menu_bar()
        
        # 상단 버튼 프레임 (그림자 효과를 위한 컨테이너)
        top_container = tk.Frame(self.root, bg='#e9ecef', height=65)
        top_container.pack(side=tk.TOP, fill=tk.X)
        top_container.pack_propagate(False)
        
        # 상단 버튼 프레임
        top_frame = tk.Frame(top_container, bg='#ffffff', relief='flat', bd=0)
        top_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # 1. 파일 관리 그룹 (PDF 열기, 저장, 새 창)
        file_frame = tk.Frame(top_frame, bg='#ffffff')
        file_frame.pack(side=tk.LEFT, padx=8)
        
        ModernButton(file_frame, text="PDF 열기", command=self.open_pdf, 
                    bg="#2563EB", fg="white").pack(side=tk.LEFT, padx=1)
        ModernButton(file_frame, text="저장", command=self.save_pdf, 
                    bg="#1D4ED8", fg="white").pack(side=tk.LEFT, padx=1)
        ModernButton(file_frame, text="선택 저장", command=self.save_selected_pages, 
                    bg="#7C3AED", fg="white").pack(side=tk.LEFT, padx=1)
        ModernButton(file_frame, text="새 창", command=self.new_window,
                    bg="#0EA5E9", fg="white").pack(side=tk.LEFT, padx=1)
        
        # 구분선 1
        separator1 = tk.Frame(top_frame, width=1, height=32, bg="#E5E7EB")
        separator1.pack(side=tk.LEFT, padx=8)
        
        # 2. 병합 그룹 (단일 병합, 다중 병합)
        merge_frame = tk.Frame(top_frame, bg='#ffffff')
        merge_frame.pack(side=tk.LEFT, padx=4)
        
        ModernButton(merge_frame, text="병합", command=self.merge_pdf, 
                    bg="#3B82F6", fg="white").pack(side=tk.LEFT, padx=1)
        ModernButton(merge_frame, text="다중 병합", command=self.merge_multiple_pdfs,
                    bg="#0D9488", fg="white").pack(side=tk.LEFT, padx=1)
        
        # 구분선 2
        separator2 = tk.Frame(top_frame, width=1, height=32, bg="#E5E7EB")
        separator2.pack(side=tk.LEFT, padx=8)
        
        # 3. 페이지 편집 그룹 (회전, 빈페이지 삽입)
        edit_frame = tk.Frame(top_frame, bg='#ffffff')
        edit_frame.pack(side=tk.LEFT, padx=4)
        
        ModernButton(edit_frame, text="우로90°", command=lambda: self.rotate_pages(90), 
                    bg="#F59E0B", fg="white").pack(side=tk.LEFT, padx=1)
        ModernButton(edit_frame, text="좌로90°", command=lambda: self.rotate_pages(-90), 
                    bg="#D97706", fg="white").pack(side=tk.LEFT, padx=1)
        
        ModernButton(edit_frame, text="빈페이지", command=self.show_insert_blank_page_dialog, 
                    bg="#10B981", fg="white").pack(side=tk.LEFT, padx=1)
        
        # 구분선 3
        separator3 = tk.Frame(top_frame, width=1, height=32, bg="#E5E7EB")
        separator3.pack(side=tk.LEFT, padx=8)
        
        # 4. 도구 그룹 (텍스트 추출, 페이지 맞춤)
        tools_frame = tk.Frame(top_frame, bg='#ffffff')
        tools_frame.pack(side=tk.LEFT, padx=4)
        
        ModernButton(tools_frame, text="텍스트", command=self.extract_text_directly, 
                    bg="#F59E0B", fg="white").pack(side=tk.LEFT, padx=1)
        ModernButton(tools_frame, text="맞춤", command=self.fit_page_to_screen, 
                    bg="#6366F1", fg="white").pack(side=tk.LEFT, padx=1)
        
        # 정보 표시 프레임 (버튼 아래) - 모던한 카드 스타일
        info_container = tk.Frame(self.root, bg='#f8f9fa', height=40)
        info_container.pack(side=tk.TOP, fill=tk.X)
        info_container.pack_propagate(False)
        
        info_frame = tk.Frame(info_container, bg="white", relief="flat", bd=0)
        info_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=8)
        
        # 선택된 페이지 정보 (왼쪽) - 더 진한 폰트
        self.selection_info = tk.Label(info_frame, text="선택된 페이지: 없음", 
                                     bg="white", fg="#212529", font=("맑은 고딕", 11, "bold"))
        self.selection_info.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 키보드 단축키 안내 (오른쪽) - 더 진한 폰트
        shortcuts_info = tk.Label(info_frame, 
                                text="Ctrl+Z: 되돌리기 | Ctrl+Y: 다시실행 | Ctrl+G: 페이지이동 | Delete: 삭제 | Ctrl+A: 전체선택",
                                bg="white", fg="#495057", font=("맑은 고딕", 9, "bold"))
        shortcuts_info.pack(side=tk.RIGHT, padx=10, pady=5)
        
        # ── v3.3: 하단 상태표시줄 ──
        self._status_bar = tk.Frame(self.root, bg='#2d3748', height=26)
        self._status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self._status_bar.pack_propagate(False)
        
        self._status_pages = tk.Label(self._status_bar, text="페이지: 0",
                                     bg='#2d3748', fg='#e2e8f0', font=('맑은 고딕', 9))
        self._status_pages.pack(side=tk.LEFT, padx=12)
        
        self._status_filesize = tk.Label(self._status_bar, text="파일크기: -",
                                        bg='#2d3748', fg='#e2e8f0', font=('맑은 고딕', 9))
        self._status_filesize.pack(side=tk.LEFT, padx=12)
        
        self._status_zoom = tk.Label(self._status_bar, text="확대: 100%",
                                    bg='#2d3748', fg='#e2e8f0', font=('맑은 고딕', 9))
        self._status_zoom.pack(side=tk.LEFT, padx=12)
        
        self._status_undo = tk.Label(self._status_bar, text="",
                                    bg='#2d3748', fg='#cbd5e0', font=('맑은 고딕', 9))
        self._status_undo.pack(side=tk.RIGHT, padx=12)
        
        # 하단 저작권 정보 프레임 - 모던한 스타일
        copyright_container = tk.Frame(self.root, bg='#e9ecef', height=30)
        copyright_container.pack(side=tk.BOTTOM, fill=tk.X)
        copyright_container.pack_propagate(False)
        
        copyright_frame = tk.Frame(copyright_container, bg="white", relief="flat", bd=0)
        copyright_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=6)
        
        # 저작권 문구 (캐시된 사용자 정보 사용)
        copyright_text = f"© 2025 Kunhwa Engineering & Consulting {VERSION} | Developed by TaeWoong Jang | 인증된 사용자: {self._cached_user_name}"
        copyright_label = tk.Label(copyright_frame, 
                                 text=copyright_text,
                                 bg="white", fg="#495057", font=("맑은 고딕", 8, "bold"))
        copyright_label.pack(expand=True, pady=2)
        
        # 수평 분할 레이아웃 (PanedWindow) - 모던한 스타일
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 좌측 패널: 썸네일 목록
        self.setup_thumbnail_panel()
        
        # 우측 패널: 미리보기
        self.setup_preview_panel()
        
        # 패널 크기 설정
        self.paned_window.sashpos(0, 320)  # 좌측 최소 폭 320px
        
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
        
        # 패널 제목 - 파일명 강조 표시
        self.thumb_title_var = tk.StringVar(value="페이지 썸네일")
        title_container = tk.Frame(title_frame, bg='#f8f9fa')
        title_container.pack(expand=True)
        base_label = tk.Label(title_container, text="페이지 썸네일 - ", bg='#f8f9fa', fg='#212529', font=("맑은 고딕", 12, "bold"))
        base_label.pack(side=tk.LEFT)
        self.thumb_filename_label = tk.Label(title_container, text="", bg='#f8f9fa', fg='#1D4ED8', font=("맑은 고딕", 12, "bold"))
        self.thumb_filename_label.pack(side=tk.LEFT)
        
        # 썸네일 캔버스와 스크롤바 - 모던한 스타일
        self.thumb_canvas = tk.Canvas(left_frame, bg="white", highlightthickness=0, relief="flat")
        self.thumb_scrollbar = tk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.thumb_canvas.yview)
        self.thumb_scrollable_frame = tk.Frame(self.thumb_canvas, bg="white")
        
        self.thumb_canvas.create_window((0, 0), window=self.thumb_scrollable_frame, anchor="nw")
        self.thumb_canvas.configure(yscrollcommand=self.thumb_scrollbar.set)
        
        self.thumb_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20, pady=5)
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
        
        # 미리보기 캔버스와 스크롤바 - 모던한 스타일 (배경 더 진한 회색)
        self.preview_canvas = tk.Canvas(right_frame, bg="#D1D5DB", highlightthickness=0, relief="flat")
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
        # 회전 단축키: Ctrl + ] (우로 90°), Ctrl + [ (좌로 90°)
        self.root.bind_all("<Control-bracketright>", lambda e: self.rotate_pages(90))
        self.root.bind_all("<Control-bracketleft>", lambda e: self.rotate_pages(-90))
        # 새 창 단축키
        self.root.bind_all("<Control-n>", lambda e: self.new_window())
        self.root.bind_all("<Control-N>", lambda e: self.new_window())
        self.root.bind_all("<Control-Shift-N>", lambda e: self.new_window_with_file())
        # (요청) Ctrl+Alt+N 바인딩 제거
        # 복사/붙여넣기 (교차 창 지원)
        self.root.bind_all("<Control-c>", self.copy_selected_pages)
        self.root.bind_all("<Control-C>", self.copy_selected_pages)
        self.root.bind_all("<Control-v>", self.paste_pages_from_clipboard)
        self.root.bind_all("<Control-V>", self.paste_pages_from_clipboard)
        # v3.3: Undo/Redo
        self.root.bind_all("<Control-z>", lambda e: self.perform_undo())
        self.root.bind_all("<Control-Z>", lambda e: self.perform_undo())
        self.root.bind_all("<Control-y>", lambda e: self.perform_redo())
        self.root.bind_all("<Control-Y>", lambda e: self.perform_redo())
        # v3.3: Go To Page
        self.root.bind_all("<Control-g>", lambda e: self.show_goto_page_dialog())
        self.root.bind_all("<Control-G>", lambda e: self.show_goto_page_dialog())

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
        """썸네일 패널 마우스 휠 스크롤 및 줌"""
        # Ctrl 키가 눌려있으면 썸네일 크기 조정
        if event.state & 0x0004:  # Ctrl 키 상태 확인
            if event.delta > 0:
                # 휠 위로: 썸네일 크기 증가
                old_scale = self.thumb_scale
                self.thumb_scale = min(1.00, self.thumb_scale + 0.05)
                if old_scale != self.thumb_scale:
                    print(f"썸네일 크기 증가: {old_scale:.2f} → {self.thumb_scale:.2f}")
                    self.refresh_thumbnails()
            else:
                # 휠 아래로: 썸네일 크기 감소
                old_scale = self.thumb_scale
                self.thumb_scale = max(0.05, self.thumb_scale - 0.05)
                if old_scale != self.thumb_scale:
                    print(f"썸네일 크기 감소: {old_scale:.2f} → {self.thumb_scale:.2f}")
                    self.refresh_thumbnails()
        else:
            # Ctrl 키가 안 눌려있으면 일반 스크롤
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
        """PDF 파일 열기 - 개선된 버전"""
        print("PDF 열기 버튼 클릭됨 - 함수 시작")
        
        try:
            # 파일 선택 다이얼로그 - 개선된 버전
            print("파일 선택 다이얼로그 열기 시도...")
            initial_dir = os.path.expanduser("~\\Documents")  # 기본 문서 폴더
            if hasattr(self, 'last_opened_dir') and os.path.exists(self.last_opened_dir):
                initial_dir = self.last_opened_dir
                print(f"마지막 열린 디렉토리 사용: {initial_dir}")
            else:
                print(f"기본 디렉토리 사용: {initial_dir}")
            
            path = filedialog.askopenfilename(
                title="PDF 파일 선택",
                initialdir=initial_dir,
                filetypes=[
                    ("PDF Files", "*.pdf"),
                    ("All Files", "*.*")
                ]
            )
            
            print(f"선택된 파일 경로: {path}")
            
            if not path:
                print("파일이 선택되지 않음")
                return
            
            # 파일 존재 여부 확인
            if not os.path.exists(path):
                print(f"파일이 존재하지 않음: {path}")
                messagebox.showerror("오류", "선택한 파일이 존재하지 않습니다.")
                return
            
            # 파일 크기 확인 (빈 파일 방지)
            file_size = os.path.getsize(path)
            print(f"파일 크기: {file_size} bytes")
            if file_size == 0:
                print("파일이 비어있음")
                messagebox.showerror("오류", "선택한 파일이 비어있습니다.")
                return
            
            # 기존 문서가 있다면 정리
            if self.doc:
                print("기존 문서 정리 중...")
                try:
                    self.doc.close()
                    print("기존 문서 닫기 성공")
                except Exception as e:
                    print(f"기존 문서 닫기 실패: {e}")
                self.doc = None
            
            # 새 PDF 문서 열기
            print(f"PDF 문서 열기 시도: {path}")
            self.doc = fitz.open(path)
            print(f"PDF 문서 열기 성공, 페이지 수: {len(self.doc)}")
            
            # 문서 유효성 검사
            if not self.doc or len(self.doc) == 0:
                print("PDF 문서가 비어있거나 유효하지 않음")
                messagebox.showerror("오류", "PDF 파일을 읽을 수 없거나 페이지가 없습니다.")
                if self.doc:
                    self.doc.close()
                    self.doc = None
                return
            
            # 상태 초기화
            print("상태 초기화 중...")
            self.current_page_index = 0
            self.selected_indices.clear()
            
            # UI 업데이트
            print("UI 업데이트 시작...")
            self.refresh_thumbnails()
            self.update_preview()
            print("UI 업데이트 완료")
            
            # 파일명 표시 및 창 제목 업데이트
            try:
                filename = os.path.basename(path)
                print(f"파일명: {filename}")
                
                if hasattr(self, 'thumb_filename_label'):
                    self.thumb_filename_label.config(text=filename)
                    print("썸네일 파일명 라벨 업데이트 완료")
                
                if hasattr(self, 'root') and self.root:
                    self.root.title(f"Kunhwa PDF Editor {VERSION} - {filename}")
                    print("창 제목 업데이트 완료")
                
                # 성공 메시지 (선택사항)
                print(f"PDF 파일 열기 성공: {filename} ({len(self.doc)}페이지)")
                
                # 마지막으로 열린 디렉토리 저장
                try:
                    self.last_opened_dir = os.path.dirname(path)
                    print(f"마지막 열린 디렉토리 저장: {self.last_opened_dir}")
                except Exception as e:
                    print(f"디렉토리 저장 중 오류: {e}")
                
            except Exception as e:
                print(f"파일명 표시 중 오류: {e}")
                
        except Exception as e:
            error_msg = f"PDF 파일을 열 수 없습니다:\n{str(e)}"
            print(f"PDF 열기 중 예외 발생: {e}")
            print(f"예외 타입: {type(e).__name__}")
            import traceback
            print(f"상세 오류: {traceback.format_exc()}")
            
            messagebox.showerror("PDF 열기 오류", error_msg)
            print(f"PDF 열기 실패: {e}")
            
            # 오류 발생 시 문서 상태 정리
            if self.doc:
                try:
                    self.doc.close()
                except:
                    pass
                self.doc = None



    def refresh_thumbnails(self):
        """썸네일 새로고침 (v3.3: 프로그레스 바 연동 및 리팩토링)"""
        if not self.doc:
            return
        
        for widget in self.thumb_scrollable_frame.winfo_children():
            widget.destroy()
        
        self.thumbnails.clear()
        self.thumbnail_labels.clear()
        self.thumbnail_frames = []
        
        total_pages = len(self.doc)
        progress = None
        if total_pages > 30:
            progress = ProgressIndicator(self.root, title="썸네일 생성 중", maximum=total_pages)
            
        for i in range(total_pages):
            thumb, label, frame = self._create_single_thumbnail(i)
            self.thumbnails.append(thumb)
            self.thumbnail_labels.append(label)
            self.thumbnail_frames.append(frame)
            
            if progress:
                progress.update(i + 1, f"썸네일 생성 중: {i+1} / {total_pages}")
        
        if progress:
            progress.close()
            
        # 썸네일 생성 후 레이아웃 업데이트
        self.root.after(50, self.update_grid_layout)
        self.root.after(100, self.ensure_scroll_region)
        self.update_selection_highlight()
        self._update_undo_redo_menu_states()

    def _update_undo_redo_menu_states(self):
        """Undo/Redo 메뉴 상태 업데이트 (v3.3)"""
        try:
            # UndoManager의 @property can_undo/can_redo 활용
            state_undo = "normal" if self.undo_manager.can_undo else "disabled"
            state_redo = "normal" if self.undo_manager.can_redo else "disabled"
            
            if hasattr(self, 'edit_menu'):
                self.edit_menu.entryconfig("되돌리기 (Undo)", state=state_undo)
                self.edit_menu.entryconfig("다시실행 (Redo)", state=state_redo)
        except Exception as e:
            # 초기화 중에는 메뉴가 없을 수 있으므로 무시
            pass


    def _schedule_performance_optimization(self):
        """성능 최적화 스케줄링 (v3.1 호환)"""
        # 60초마다 메모리 최적화 실행
        self.root.after(60000, self._run_performance_optimization)
        
    def _run_performance_optimization(self):
        """성능 최적화 실행"""
        try:
            # 메모리 사용량 최적화
            self._optimize_memory_usage()
            
            # 다음 최적화 예약
            self.root.after(60000, self._run_performance_optimization)
            
        except Exception as e:
            print(f"성능 최적화 오류: {e}")
            # 오류 발생시에도 다음 실행 예약
            self.root.after(60000, self._run_performance_optimization)

    def _cleanup_cache(self):
        """캐시 정리 및 메모리 최적화"""
        try:
            # 캐시 크기 제한 확인
            if len(self._thumbnail_cache) > self._cache_size_limit:
                # 가장 오래된 캐시 항목 제거
                oldest_keys = sorted(self._thumbnail_cache.keys(), key=lambda k: self._thumbnail_cache[k].get('timestamp', 0))[:20]
                for key in oldest_keys:
                    del self._thumbnail_cache[key]
            
            if len(self._preview_cache) > self._cache_size_limit:
                oldest_keys = sorted(self._preview_cache.keys(), key=lambda k: self._preview_cache[k].get('timestamp', 0))[:20]
                for key in oldest_keys:
                    del self._preview_cache[key]
                    
        except Exception as e:
            print(f"캐시 정리 오류: {e}")



    def _create_single_thumbnail(self, index):
        """단일 썸네일 생성 (최적화)"""
        try:
            # 캐시 확인
            cache_key = f"thumb_{index}_{self.thumb_scale}"
            if cache_key in self._thumbnail_cache:
                self._cache_hits += 1
                cached_data = self._thumbnail_cache[cache_key]
                return self._create_thumbnail_widget(index, cached_data['image'])
            
            self._cache_misses += 1
            
            # 페이지 정보 가져오기
            page = self.doc[index]
            page_rect = page.rect
            
            # 최적화된 스케일 계산
            scale_factor = self._calculate_optimal_scale(page_rect)
            
            # 썸네일 이미지 생성
            pix = page.get_pixmap(matrix=fitz.Matrix(scale_factor, scale_factor))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            thumb = ImageTk.PhotoImage(img)
            
            # 캐시에 저장
            self._thumbnail_cache[cache_key] = {
                'image': thumb,
                'timestamp': time.time()
            }
            
            # 위젯 생성 및 반환
            return self._create_thumbnail_widget(index, thumb)
            
        except Exception as e:
            print(f"썸네일 생성 오류 (페이지 {index}): {e}")
            return None, None, None

    def _calculate_optimal_scale(self, page_rect):
        """최적 스케일 계산"""
        base_thumb_width = 110
        base_thumb_height = 150
        
        max_thumb_width = base_thumb_width * self.thumb_scale / 0.20
        max_thumb_height = base_thumb_height * self.thumb_scale / 0.20
        
        if page_rect.width > page_rect.height:  # 가로형
            scale_factor = min(max_thumb_width / page_rect.width, max_thumb_height / page_rect.height)
        else:  # 세로형
            scale_factor = min(max_thumb_height / page_rect.height, max_thumb_width / page_rect.width)
        
        return max(scale_factor, 0.05)

    def _create_thumbnail_widget(self, index, thumb_image):
        """썸네일 위젯 생성"""
        # 프레임 생성
        frame = tk.Frame(self.thumb_scrollable_frame, bg="white", relief="flat", bd=0)
        
        # 썸네일 라벨
        label = tk.Label(frame, image=thumb_image, borderwidth=2, relief="solid", bg="white")
        label.bind("<Button-1>", partial(self.handle_selection, index=index))
        label.bind("<Double-Button-1>", partial(self.on_double_click, index=index))
        label.bind("<B1-Motion>", partial(self.on_drag, index=index))
        label.bind("<ButtonRelease-1>", partial(self.on_drop, index=index))
        label.bind("<Enter>", lambda e, idx=index: self.on_enter(e, idx))
        label.bind("<Leave>", lambda e, idx=index: self.on_leave(e, idx))
        label.bind("<MouseWheel>", self.on_thumb_mousewheel)
        
        # 드래그 앤 드롭 이벤트
        if DRAG_DROP_AVAILABLE:
            try:
                label.drop_target_register(DND_FILES, DND_TEXT)
                label.dnd_bind('<<Drop>>', self.on_generic_drop)
                label.drag_source_register(1, DND_TEXT)
                label.dnd_bind('<<DragInitCmd>>', lambda e, idx=index: ("copy", DND_TEXT, f"PDFTHUMB::{self.window_id}::{idx}"))
            except Exception:
                pass
        
        label.pack(padx=5, pady=5)
        
        # 페이지 번호 라벨
        num_label = tk.Label(frame, text=f"{index+1}", font=("맑은 고딕", 10, "bold"), 
                           bg="white", fg="#212529")
        num_label.pack(pady=(0, 5))
        num_label.bind("<MouseWheel>", self.on_thumb_mousewheel)
        
        if DRAG_DROP_AVAILABLE:
            try:
                num_label.drop_target_register(DND_FILES, DND_TEXT)
                num_label.dnd_bind('<<Drop>>', self.on_generic_drop)
                num_label.drag_source_register(1, DND_TEXT)
                num_label.dnd_bind('<<DragInitCmd>>', lambda e, idx=index: ("copy", DND_TEXT, f"PDFTHUMB::{self.window_id}::{idx}"))
            except Exception:
                pass
        
        # 리스트에 추가 (호출하는 곳에서 수행하도록 변경)
        # self.thumbnails.append(thumb_image)
        # self.thumbnail_labels.append(label)
        # self.thumbnail_frames.append(frame)
        
        # 프레임을 그리드에 배치
        frame.pack(padx=5, pady=5)
        
        return thumb_image, label, frame

    def _optimize_memory_usage(self):
        """메모리 사용량 최적화 (v3.1 호환)"""
        try:
            # 가비지 컬렉션 강제 실행
            import gc
            gc.collect()
            
            # 캐시 크기 조정
            if len(self._thumbnail_cache) > self._cache_size_limit * 0.8:
                self._cleanup_cache()
                
        except Exception as e:
            print(f"메모리 최적화 오류: {e}")

    def _force_cleanup(self):
        """강제 메모리 정리"""
        try:
            # 모든 캐시 비우기
            self._thumbnail_cache.clear()
            self._preview_cache.clear()
            
            # 가비지 컬렉션 강제 실행
            import gc
            gc.collect()
            
            print("강제 메모리 정리 완료")
            
        except Exception as e:
            print(f"강제 정리 오류: {e}")

    def _get_cache_stats(self):
        """캐시 통계 정보 반환"""
        return {
            'thumbnail_cache_size': len(self._thumbnail_cache),
            'preview_cache_size': len(self._preview_cache),
            'cache_hits': self._cache_hits,
            'cache_misses': self._cache_misses,
            'hit_rate': self._cache_hits / (self._cache_hits + self._cache_misses) if (self._cache_hits + self._cache_misses) > 0 else 0
        }

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
        """썸네일을 창 크기에 맞춰 자동으로 배열하는 그리드 레이아웃 (v2.4 최적화)"""
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
            thumb_width = sample_frame.winfo_reqwidth() + 20  # 좌우 패딩 10px씩
            thumb_height = sample_frame.winfo_reqheight() + 20  # 상하 패딩 10px씩
            
            # 스크롤바 너비와 여백 고려
            scrollbar_width = 20  # 스크롤바 너비
            total_margin = 40  # 좌우 여백 20px씩 (적절한 여백)
            
            # 사용 가능한 너비 계산 (스크롤바와 여백 제외)
            available_width = canvas_width - scrollbar_width - total_margin
            
            # 안전한 열 수 계산 (썸네일이 절대 잘리지 않도록)
            # 레이아웃 크기에 따라 최적의 열 수 자동 계산
            col_count = max(1, available_width // thumb_width)
            
            # 너무 많은 열은 방지 (UI가 복잡해지지 않도록)
            max_cols = 3  # 최대 3열까지만 허용
            if col_count > max_cols:
                col_count = max_cols
                print(f"최대 {max_cols}열로 제한")
            
            # 실제 그리드가 캔버스 너비를 초과하지 않는지 최종 확인
            actual_grid_width = col_count * thumb_width + total_margin
            if actual_grid_width > canvas_width:
                col_count = max(1, col_count - 1)
                print(f"공간 부족으로 {col_count}열로 조정")
            
            # 행 수 계산
            total_pages = len(self.thumbnail_frames)
            row_count = (total_pages + col_count - 1) // col_count  # 올림 나눗셈
            
            # 그리드 배치
            for i, frame in enumerate(self.thumbnail_frames):
                row = i // col_count
                col = i % col_count
                
                # 그리드 위치 설정
                frame.grid(row=row, column=col, padx=10, pady=10, sticky="nw")
                
                # 프레임이 잘리지 않도록 확인
                frame.grid_propagate(False)
            
            # 스크롤 영역 갱신
            self.thumb_canvas.update_idletasks()
            
            # 전체 그리드 크기 계산 (실제 사용된 공간만큼)
            total_width = col_count * thumb_width + total_margin
            total_height = row_count * thumb_height + 40  # 상하 여백 포함
            
            # 스크롤 영역 설정 - 모든 썸네일이 보이도록 충분한 공간 확보
            self.thumb_canvas.config(scrollregion=(0, 0, total_width, total_height))
            
            print(f"그리드 레이아웃 업데이트: {col_count}열 x {row_count}행, 총 {total_pages}페이지")
            print(f"캔버스 크기: {canvas_width}x{canvas_height}, 그리드 크기: {total_width}x{total_height}")
            print(f"사용 가능한 너비: {available_width}, 썸네일 너비: {thumb_width}")
            print(f"여백: {total_margin}px, 스크롤바: {scrollbar_width}px")
            print(f"레이아웃 최적화: {col_count}열로 설정 (공간 효율성: {col_count * thumb_width / available_width:.1%})")

    def ensure_scroll_region(self):
        """스크롤 영역이 제대로 설정되었는지 확인하고 필요시 수정 (v2.4 최적화)"""
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
            
            # 스크롤 영역 계산 - 실제 그리드 크기에 맞춤
            scroll_width = max(current_scroll[2] + 40, canvas_width)  # 좌우 여백 20px씩
            scroll_height = max(current_scroll[3] + 40, canvas_height)  # 상하 여백 20px씩
            
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
    
    def on_interwindow_drop(self, event):
        """다른 창에서 넘어온 썸네일 드롭 처리 (텍스트 페이로드 기반)"""
        try:
            data = getattr(event, 'data', '') or ''
            if not data or 'PDFTHUMB::' not in data:
                return
            # 예: 'PDFTHUMB::12'
            try:
                payload = data.split('PDFTHUMB::')[-1]
                src_win, src_index_str = payload.split('::')
                src_index = int(src_index_str)
            except Exception:
                return
            
            # 소스 창 찾기: 현재 포커스된 창이 아닌 다른 에디터 중 하나로 가정
            src_app = None
            for app in OPEN_EDITORS:
                if app is not self and getattr(app, 'window_id', None) == src_win and app.doc is not None and src_index < len(app.doc):
                    src_app = app
                    break
            if src_app is None:
                return
            
            # 드롭 타겟 위치 계산 (없으면 마지막)
            insert_pos = len(self.doc) if self.doc else 0
            if hasattr(self, 'drop_target_index'):
                insert_pos = max(0, int(self.drop_target_index))
                delattr(self, 'drop_target_index')
            
            # 문서가 없으면 새 문서 생성
            if not self.doc:
                self.doc = fitz.open()
                self.current_page_index = 0
            
            # PDF 객체 단위 안전복사: 페이지를 통째로 import
            src_doc = src_app.doc
            # 임시 문서에 해당 페이지만 추출 후 대상에 삽입
            temp = fitz.open()
            temp.insert_pdf(src_doc, from_page=src_index, to_page=src_index)
            # 대상 문서에 원하는 위치로 삽입
            self.doc.insert_pdf(temp, from_page=0, to_page=0, start_at=insert_pos)
            temp.close()
            
            self.refresh_thumbnails()
            self.update_preview()
            # 이동: 소스 창에서 원본 페이지 삭제 (객체 보존 상태 유지)
            try:
                src_doc.delete_page(src_index)
                src_app.refresh_thumbnails()
                src_app.update_preview()
            except Exception as e:
                print(f"소스 페이지 삭제 실패(복사로 처리): {e}")
        except Exception as e:
            print(f"교차 창 드롭 처리 오류: {e}")
    
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
            
        # Undo 상태 저장 (v3.3)
        if self.doc:
            self.undo_manager.save_state(self.doc)

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
        
        # v3.3: Undo 상태 저장
        self.undo_manager.save_state(self.doc, "페이지 삭제")
        
        for idx in sorted(self.selected_indices, reverse=True):
            self.doc.delete_page(idx)
        
        max_idx = len(self.doc) - 1
        self.selected_indices = {min(i, max_idx) for i in self.selected_indices if i <= max_idx}
        
        # 현재 페이지 인덱스 조정
        if self.current_page_index >= len(self.doc):
            self.current_page_index = max(0, len(self.doc) - 1)
        
        self.refresh_thumbnails()
        self.update_preview()
        self._update_status_bar()

    def rotate_pages(self, angle):
        """페이지 회전 (기존 함수 수정)"""
        if not self.selected_indices:
            messagebox.showwarning("경고", "회전할 페이지를 선택해주세요.")
            return
        
        # v3.3: Undo 상태 저장
        self.undo_manager.save_state(self.doc, "페이지 회전")
        
        for idx in self.selected_indices:
            current_rotation = self.doc[idx].rotation
            new_rotation = (current_rotation + angle) % 360
            self.doc[idx].set_rotation(new_rotation)
        
        self.refresh_thumbnails()
        self.update_preview()
        self._update_status_bar()

    def move_pages(self, direction):
        if not self.selected_indices:
            return
        
        # v3.3: Undo 상태 저장
        self.undo_manager.save_state(self.doc, "페이지 이동")
        
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
        self._update_status_bar()

    def save_pdf(self):
        # 기본 폴더: 현재 열린 파일의 폴더
        initialdir = None
        try:
            if getattr(self, 'doc', None) and getattr(self.doc, 'name', None):
                initialdir = os.path.dirname(self.doc.name)
        except Exception:
            initialdir = None
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialdir=initialdir)
        if path:
            self.doc.save(path)
            messagebox.showinfo("저장 완료", f"{path}로 저장됨")

    def save_selected_pages(self):
        """선택한 페이지를 PDF 또는 JPG로 저장"""
        if not self.doc:
            messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
            return
        
        # 선택된 페이지가 없으면 현재 페이지 사용
        if not self.selected_indices:
            selected_pages = [self.current_page_index]
        else:
            selected_pages = sorted(list(self.selected_indices))
        
        # 저장 형식 선택 다이얼로그
        self.show_save_format_dialog(selected_pages)
    
    def show_save_format_dialog(self, selected_pages):
        """저장 형식 선택 다이얼로그 - 통일된 UI"""
        dialog = Toplevel(self.root)
        dialog.title("선택 페이지 저장")
        dialog.geometry("450x350")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#f8f9fa')
        
        # 중앙 정렬
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 100))
        
        # 메인 컨테이너
        main_frame = tk.Frame(dialog, bg='#f8f9fa')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 제목
        title_label = tk.Label(main_frame, text=f"선택된 {len(selected_pages)}개 페이지 저장", 
                              font=("맑은 고딕", 16, "bold"), bg='#f8f9fa', fg="#1f2937")
        title_label.pack(pady=(0, 15))
        
        # 페이지 목록 표시
        if len(selected_pages) <= 10:
            page_list = ", ".join([str(p + 1) for p in selected_pages])
        else:
            page_list = f"{selected_pages[0] + 1}, {selected_pages[1] + 1}, ..., {selected_pages[-1] + 1}"
        
        info_label = tk.Label(main_frame, text=f"페이지: {page_list}", 
                             font=("맑은 고딕", 11), fg="#6B7280", bg='#f8f9fa')
        info_label.pack(pady=(0, 25))
        
        # 저장 형식 버튼들
        button_frame = tk.Frame(main_frame, bg='#f8f9fa')
        button_frame.pack(pady=10)
        
        # 버튼 스타일 통일
        button_width = 25
        button_height = 2
        button_font = ("맑은 고딕", 11, "bold")
        
        # PDF로 저장 버튼
        pdf_btn = ModernButton(button_frame, text="📄 PDF로 저장", 
                              command=lambda: self.save_pages_as_pdf(selected_pages, dialog),
                              bg="#DC2626", fg="white", font=button_font,
                              width=button_width, height=button_height)
        pdf_btn.pack(pady=8, fill=tk.X, padx=30)
        
        # JPG로 저장 버튼
        jpg_btn = ModernButton(button_frame, text="🖼️ JPG로 저장", 
                              command=lambda: self.save_pages_as_jpg(selected_pages, dialog),
                              bg="#059669", fg="white", font=button_font,
                              width=button_width, height=button_height)
        jpg_btn.pack(pady=8, fill=tk.X, padx=30)
        
        # 구분선
        separator = tk.Frame(button_frame, height=2, bg="#E5E7EB")
        separator.pack(fill=tk.X, pady=15, padx=30)
        
        # 취소 버튼
        cancel_btn = ModernButton(button_frame, text="취소", 
                                 command=dialog.destroy,
                                 bg="#6B7280", fg="white", font=button_font,
                                 width=button_width, height=button_height)
        cancel_btn.pack(pady=8, fill=tk.X, padx=30)
    
    def save_pages_as_pdf(self, selected_pages, dialog):
        """선택된 페이지들을 PDF로 저장 (원본 화질 유지)"""
        try:
            # 저장 경로 선택
            # 기본 폴더: 현재 열린 파일의 폴더
            initialdir = None
            try:
                if getattr(self, 'doc', None) and getattr(self.doc, 'name', None):
                    initialdir = os.path.dirname(self.doc.name)
            except Exception:
                initialdir = None
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
                title="PDF로 저장",
                initialdir=initialdir
            )
            
            if not file_path:
                return
            
            # 새 PDF 문서 생성
            new_doc = fitz.open()
            
            # 선택된 페이지들을 새 문서에 복사 (원본 화질 유지)
            for page_idx in selected_pages:
                try:
                    # 원본 페이지를 그대로 복사 (화질 손실 없음)
                    new_doc.insert_pdf(self.doc, from_page=page_idx, to_page=page_idx)
                    
                except Exception as e:
                    print(f"페이지 {page_idx + 1} 복사 중 오류: {e}")
                    continue
            
            # PDF 저장 (원본 화질 유지)
            new_doc.save(file_path, garbage=4, deflate=True)
            new_doc.close()
            
            dialog.destroy()
            messagebox.showinfo("저장 완료", 
                              f"{len(selected_pages)}개 페이지가 원본 화질로 PDF 저장되었습니다.\n경로: {file_path}")
            
        except Exception as e:
            messagebox.showerror("오류", f"PDF 저장 중 오류가 발생했습니다: {str(e)}")
    
    def save_pages_as_jpg(self, selected_pages, dialog):
        """선택된 페이지들을 JPG로 저장 (파일명 입력 기능)"""
        try:
            # 저장 폴더 선택 (기본: 현재 열린 파일 폴더)
            initialdir = None
            try:
                if getattr(self, 'doc', None) and getattr(self.doc, 'name', None):
                    initialdir = os.path.dirname(self.doc.name)
            except Exception:
                initialdir = None
            folder_path = filedialog.askdirectory(title="JPG 파일들을 저장할 폴더 선택", initialdir=initialdir)
            
            if not folder_path:
                return
            
            # 파일명 입력 받기
            filename_dialog = Toplevel(dialog)
            filename_dialog.title("JPG 파일명 입력")
            filename_dialog.geometry("450x280")
            filename_dialog.transient(dialog)
            filename_dialog.grab_set()
            filename_dialog.configure(bg='#f8f9fa')
            
            # 중앙 정렬
            filename_dialog.geometry("+%d+%d" % (dialog.winfo_rootx() + 50, dialog.winfo_rooty() + 50))
            
            # 메인 컨테이너
            main_frame = tk.Frame(filename_dialog, bg='#f8f9fa')
            main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # 안내 메시지
            tk.Label(main_frame, text="JPG 파일명을 입력하세요", 
                    font=("맑은 고딕", 14, "bold"), bg='#f8f9fa', fg="#1f2937").pack(pady=(0, 10))
            
            if len(selected_pages) == 1:
                info_text = "단일 파일로 저장됩니다"
            else:
                info_text = f"{len(selected_pages)}개 파일로 저장됩니다\n(파일명_001.jpg, 파일명_002.jpg, ...)"
            
            tk.Label(main_frame, text=info_text, 
                    font=("맑은 고딕", 10), fg="#6B7280", bg='#f8f9fa').pack(pady=(0, 20))
            
            # 파일명 입력 프레임
            input_frame = tk.Frame(main_frame, bg='#f8f9fa')
            input_frame.pack(pady=15, padx=20, fill=tk.X)
            
            tk.Label(input_frame, text="파일명:", font=("맑은 고딕", 11, "bold"), 
                    bg='#f8f9fa', fg="#374151").pack(side=tk.LEFT)
            filename_entry = tk.Entry(input_frame, font=("맑은 고딕", 11), width=20)
            filename_entry.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True)
            tk.Label(input_frame, text=".jpg", font=("맑은 고딕", 11), 
                    bg='#f8f9fa', fg="#6B7280").pack(side=tk.LEFT)
            
            # 기본값 설정
            if len(selected_pages) == 1:
                default_name = f"page_{selected_pages[0] + 1}"
            else:
                default_name = "pages"
            filename_entry.insert(0, default_name)
            filename_entry.select_range(0, tk.END)
            filename_entry.focus()
            
            # 버튼 프레임
            button_frame = tk.Frame(main_frame, bg='#f8f9fa')
            button_frame.pack(pady=25)
            
            # 버튼 스타일 통일
            button_width = 12
            button_height = 1
            button_font = ("맑은 고딕", 10, "bold")
            
            # 저장 버튼
            save_btn = ModernButton(button_frame, text="저장", 
                                   command=lambda: self.process_jpg_save(selected_pages, folder_path, 
                                                                       filename_entry.get().strip(), 
                                                                       dialog, filename_dialog),
                                   bg="#059669", fg="white", font=button_font,
                                   width=button_width, height=button_height)
            save_btn.pack(side=tk.LEFT, padx=(0, 15))
            
            # 취소 버튼
            cancel_btn = ModernButton(button_frame, text="취소", 
                                     command=filename_dialog.destroy,
                                     bg="#6B7280", fg="white")
            cancel_btn.pack(side=tk.LEFT)
            
            # Enter 키로 저장
            filename_entry.bind('<Return>', lambda e: self.process_jpg_save(selected_pages, folder_path, 
                                                                          filename_entry.get().strip(), 
                                                                          dialog, filename_dialog))
            
        except Exception as e:
            messagebox.showerror("오류", f"JPG 저장 준비 중 오류가 발생했습니다: {str(e)}")
    
    def process_jpg_save(self, selected_pages, folder_path, base_filename, main_dialog, filename_dialog):
        """JPG 저장 처리"""
        try:
            if not base_filename:
                messagebox.showwarning("경고", "파일명을 입력해주세요.")
                return
            
            # 파일명에서 특수문자 제거
            import re
            base_filename = re.sub(r'[<>:"/\\|?*]', '_', base_filename)
            
            # 각 페이지를 JPG로 저장
            saved_files = []
            
            for i, page_idx in enumerate(selected_pages):
                try:
                    # 페이지를 이미지로 렌더링
                    page = self.doc[page_idx]
                    
                    # 고해상도로 렌더링 (300 DPI)
                    zoom_factor = 300 / 72  # 72 DPI -> 300 DPI
                    matrix = fitz.Matrix(zoom_factor, zoom_factor)
                    pix = page.get_pixmap(matrix=matrix)
                    
                    # PIL Image로 변환
                    img_data = pix.samples
                    img = Image.frombytes("RGB", [pix.width, pix.height], img_data)
                    
                    # 파일명 생성
                    if len(selected_pages) == 1:
                        filename = f"{base_filename}.jpg"
                    else:
                        filename = f"{base_filename}_{i+1:03d}.jpg"
                    
                    file_path = os.path.join(folder_path, filename)
                    
                    # 파일이 이미 존재하는지 확인
                    if os.path.exists(file_path):
                        result = messagebox.askyesno("파일 덮어쓰기", 
                                                   f"'{filename}' 파일이 이미 존재합니다.\n덮어쓰시겠습니까?")
                        if not result:
                            continue
                    
                    # JPG로 저장 (고품질)
                    img.save(file_path, "JPEG", quality=95, optimize=True)
                    saved_files.append(filename)
                    
                except Exception as e:
                    print(f"페이지 {page_idx + 1} JPG 저장 중 오류: {e}")
                    continue
            
            # 다이얼로그 닫기
            filename_dialog.destroy()
            main_dialog.destroy()
            
            if saved_files:
                messagebox.showinfo("저장 완료", 
                                  f"{len(saved_files)}개 페이지가 고화질 JPG로 저장되었습니다.\n"
                                  f"폴더: {folder_path}\n"
                                  f"파일: {saved_files[0]}" + (f" 외 {len(saved_files)-1}개" if len(saved_files) > 1 else ""))
            else:
                messagebox.showerror("오류", "저장된 파일이 없습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"JPG 저장 중 오류가 발생했습니다: {str(e)}")
    
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
        if self.doc:
            self.undo_manager.save_state(self.doc)
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
            self.doc.new_page(width=width_pt, height=height_pt)
            
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
            
        # Undo 상태 저장 (v3.3)
        self.undo_manager.save_state(self.doc)
        
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

    def merge_multiple_pdfs(self):
        """여러 PDF 파일을 사용자가 지정한 순서로 병합하여 새로운 PDF 생성"""
        # 파일 다중 선택
        file_paths = filedialog.askopenfilenames(
            title="병합할 PDF 파일 선택 (여러 개)",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not file_paths:
            return

        # 기본 정렬: 파일명 오름차순
        sorted_paths = sorted(file_paths, key=lambda p: os.path.basename(p))

        # 순서 조정 다이얼로그 열기
        ordered_paths = self._open_order_dialog(sorted_paths)
        if not ordered_paths:
            return

        # Undo 상태 저장 (v3.3)
        if self.doc:
            self.undo_manager.save_state(self.doc)

        # 새로운 PDF 문서 생성
        try:
            # 원본 화질/사이즈 보장 안내
            print("=== PDF 병합 시작 ===")
            print("원본 화질과 사이즈를 100% 유지하여 병합합니다.")
            print("모든 페이지가 원본과 동일한 품질로 병합됩니다.")
            print("=" * 30)
            
            merged_doc = fitz.open()
            total_pages = 0
            
            # 선택된 파일들을 순서대로 병합
            for path in ordered_paths:
                try:
                    print(f"병합 중: {os.path.basename(path)}")
                    source_doc = fitz.open(path)
                    
                    if not source_doc or len(source_doc) == 0:
                        print(f"빈 파일 또는 읽을 수 없는 파일: {path}")
                        continue
                    
                    # 모든 페이지를 새 문서에 추가 (원본 화질/사이즈 100% 유지)
                    for page_num in range(len(source_doc)):
                        try:
                            source_page = source_doc[page_num]
                            
                            # 원본 페이지의 모든 정보를 그대로 복사
                            new_page = merged_doc.new_page(
                                width=source_page.rect.width, 
                                height=source_page.rect.height
                            )
                            
                            # 원본 페이지를 그대로 복사 (화질 손실 없음, 사이즈 그대로)
                            new_page.show_pdf_page(new_page.rect, source_doc, page_num)
                            
                            total_pages += 1
                            print(f"페이지 {page_num + 1} 복사 완료 (원본 화질/사이즈 100% 유지)")
                            
                        except Exception as e:
                            print(f"페이지 {page_num} 복사 실패: {e}")
                            continue
                    
                    source_doc.close()
                    print(f"'{os.path.basename(path)}' 병합 완료 ({len(source_doc)}페이지)")
                    
                except Exception as e:
                    print(f"'{os.path.basename(path)}' 병합 실패: {e}")
                    continue
            
            if total_pages == 0:
                messagebox.showwarning("경고", "병합할 수 있는 페이지가 없습니다.")
                merged_doc.close()
                return
            
            # 병합된 PDF를 새 파일로 저장
            save_path = filedialog.asksaveasfilename(
                title="병합된 PDF 저장",
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                initialfile="병합된_문서.pdf"
            )
            
            if save_path:
                try:
                    merged_doc.save(save_path)
                    merged_doc.close()
                    
                    # 성공 메시지
                    messagebox.showinfo("완료", 
                        f"PDF 병합 완료!\n\n"
                        f"파일: {os.path.basename(save_path)}\n"
                        f"총 페이지: {total_pages}페이지\n"
                        f"화질: 원본 100% 유지\n"
                        f"사이즈: 원본 그대로 유지\n"
                        f"저장 위치: {save_path}")
                    
                    # 저장 후 새로 생성된 PDF를 열기 옵션 제공
                    if messagebox.askyesno("확인", "병합된 PDF를 지금 열까요?"):
                        self.open_pdf_from_path(save_path)
                        
                except Exception as e:
                    messagebox.showerror("저장 오류", f"PDF 저장 중 오류가 발생했습니다:\n{str(e)}")
                    print(f"PDF 저장 실패: {e}")
            else:
                merged_doc.close()
                
        except Exception as e:
            error_msg = f"PDF 병합 중 오류가 발생했습니다:\n{str(e)}"
            messagebox.showerror("병합 오류", error_msg)
            print(f"PDF 병합 실패: {e}")

    def _insert_pdf_all_pages_at(self, file_path, insert_pos):
        """file_path의 모든 페이지를 insert_pos 위치부터 순서대로 삽입. 성공적으로 추가한 페이지 수 반환"""
        merge_doc = None
        try:
            merge_doc = fitz.open(file_path)
            if not merge_doc:
                return 0

            added_indices = []
            # 우선 뒤에 페이지를 생성하여 복사한 뒤, 원하는 위치로 이동
            for i in range(len(merge_doc)):
                try:
                    src_page = merge_doc[i]
                    new_page = self.doc.new_page(width=src_page.rect.width, height=src_page.rect.height)
                    new_page.insert_image(new_page.rect, pixmap=src_page.get_pixmap())
                    added_indices.append(len(self.doc) - 1)
                except Exception as e:
                    print(f"페이지 복사 실패({i}): {e}")
                    continue

            # 생성된 페이지들을 원하는 위치로 순서 유지하며 이동
            moved = 0
            for i, page_index in enumerate(added_indices):
                try:
                    self.doc.move_page(page_index, insert_pos + i)
                    moved += 1
                except Exception as e:
                    print(f"페이지 이동 실패: {e}")
                    continue

            return moved
        finally:
            try:
                if merge_doc:
                    merge_doc.close()
            except Exception:
                pass

    def _open_order_dialog(self, paths):
        """파일 순서를 사용자에게 확인/수정받는 간단한 다이얼로그. 최종 순서 리스트 반환"""
        dialog = Toplevel(self.root)
        dialog.title("다중 병합 - 파일 순서 정하기")
        dialog.geometry("520x420")
        dialog.transient(self.root)
        dialog.grab_set()

        # 내부 상태: 현재 경로 순서
        current_paths = list(paths)

        # 리스트박스와 스크롤바
        frame = tk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
        listbox = tk.Listbox(frame, selectmode=tk.EXTENDED, yscrollcommand=scrollbar.set, height=14)
        scrollbar.config(command=listbox.yview)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def rebuild_listbox():
            listbox.delete(0, tk.END)
            for p in current_paths:
                listbox.insert(tk.END, os.path.basename(p))

        # 초기 아이템 추가
        rebuild_listbox()

        # 버튼 영역
        btns = tk.Frame(dialog)
        btns.pack(fill=tk.X, padx=10, pady=6)

        def move_up():
            sel = list(listbox.curselection())
            if not sel:
                return
            if 0 in sel:
                return
            # 위로 이동: 선택된 인덱스 순서대로 위 요소와 교환
            for idx in sel:
                current_paths[idx-1], current_paths[idx] = current_paths[idx], current_paths[idx-1]
            rebuild_listbox()
            # 선택 재설정
            listbox.selection_clear(0, tk.END)
            new_sel = [s - 1 for s in sel]
            for idx in new_sel:
                listbox.selection_set(idx)
            listbox.see(min(new_sel))

        def move_down():
            sel = list(listbox.curselection())
            if not sel:
                return
            if sel[-1] == listbox.size() - 1:
                return
            # 아래로 이동: 선택된 인덱스를 역순으로 아래 요소와 교환
            for idx in reversed(sel):
                current_paths[idx+1], current_paths[idx] = current_paths[idx], current_paths[idx+1]
            rebuild_listbox()
            # 선택 재설정
            listbox.selection_clear(0, tk.END)
            new_sel = [s + 1 for s in sel]
            for idx in new_sel:
                listbox.selection_set(idx)
            listbox.see(max(new_sel))

        def sort_by_name():
            # 파일명 기준 오름차순 정렬
            current_paths.sort(key=lambda p: os.path.basename(p))
            rebuild_listbox()

        left = tk.Frame(btns)
        left.pack(side=tk.LEFT)
        ModernButton(left, text="위로", command=move_up, bg="#64748B", fg="white").pack(side=tk.LEFT, padx=2)
        ModernButton(left, text="아래로", command=move_down, bg="#475569", fg="white").pack(side=tk.LEFT, padx=2)
        ModernButton(left, text="이름 정렬", command=sort_by_name, bg="#6B7280", fg="white").pack(side=tk.LEFT, padx=6)

        right = tk.Frame(btns)
        right.pack(side=tk.RIGHT)
        result = {"ok": False}

        def on_ok():
            result["ok"] = True
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        ModernButton(right, text="확인", command=on_ok, bg="#2563EB", fg="white").pack(side=tk.LEFT, padx=4)
        ModernButton(right, text="취소", command=on_cancel, bg="#EF4444", fg="white").pack(side=tk.LEFT, padx=4)

        # 중앙 배치
        try:
            dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 120, self.root.winfo_rooty() + 120))
        except Exception:
            pass

        dialog.wait_window()

        if not result["ok"]:
            return None

        # 현재 경로 순서를 반환
        return list(current_paths)

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

    def set_performance_mode(self, mode):
        """성능 모드 설정 (새로 추가)"""
        old_scale = self.thumb_scale
        
        if mode == "high":
            self.thumb_scale = 0.12  # 저해상도, 빠른 처리
            self.performance_mode = "high"
            self.current_mode_label.set("현재: 🚀 고성능 모드")
            messagebox.showinfo("성능 모드", 
                              "🚀 고성능 모드로 설정되었습니다.\n\n"
                              "• 빠른 처리 속도를 위해 썸네일 해상도가 낮아집니다\n"
                              "• 대용량 PDF 파일 처리에 최적화\n"
                              "• 메모리 사용량 감소\n"
                              "• 권장: 50MB 이상의 PDF 파일")
            
        elif mode == "balanced":
            self.thumb_scale = 0.20  # 기본 설정
            self.performance_mode = "balanced"
            self.current_mode_label.set("현재: ⚖️ 균형 모드")
            messagebox.showinfo("성능 모드", 
                              "⚖️ 균형 모드로 설정되었습니다.\n\n"
                              "• 속도와 품질의 균형\n"
                              "• 일반적인 용도에 적합\n"
                              "• 권장: 10-50MB PDF 파일")
            
        elif mode == "quality":
            self.thumb_scale = 0.30  # 고해상도, 느린 처리
            self.performance_mode = "quality"
            self.current_mode_label.set("현재: 🎨 고품질 모드")
            messagebox.showinfo("성능 모드", 
                              "🎨 고품질 모드로 설정되었습니다.\n\n"
                              "• 높은 품질의 썸네일\n"
                              "• 처리 속도가 느려질 수 있음\n"
                              "• 권장: 10MB 미만의 PDF 파일")
        
        # 설정이 변경된 경우에만 썸네일 새로고침
        if old_scale != self.thumb_scale:
            print(f"성능 모드 변경: {old_scale:.2f} → {self.thumb_scale:.2f}")
            if self.doc:
                self.refresh_thumbnails()

    def show_users_list(self):
        """현재 등록된 사용자 목록 표시"""
        # 관리자 권한 확인
        if not self._check_admin_permission():
            return
        
        users_data = load_encrypted_users()
        if not users_data:
            messagebox.showwarning("경고", "사용자 정보를 불러올 수 없습니다.")
            return
        
        dialog = Toplevel(self.root)
        dialog.title("등록된 사용자 목록")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#f8f9fa')
        
        # 중앙 정렬
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 100))
        
        # 메인 컨테이너
        main_frame = tk.Frame(dialog, bg='#f8f9fa')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 제목
        title_label = tk.Label(main_frame, text="등록된 사용자 목록", 
                              font=("맑은 고딕", 16, "bold"), bg='#f8f9fa', fg="#1f2937")
        title_label.pack(pady=(0, 10))
        
        # 총 사용자 수 (제목 바로 아래 표시)
        info_text = f"총 {len(users_data['users'])}명의 사용자가 등록되어 있습니다."
        info_label = tk.Label(main_frame, text=info_text, 
                             font=("맑은 고딕", 10), fg="#6B7280", bg='#f8f9fa')
        info_label.pack(pady=(0, 12))

        # 목록 영역 컨테이너 (리스트/스크롤 묶음)
        list_frame = tk.Frame(main_frame, bg='#f8f9fa')
        list_frame.pack(fill=tk.BOTH, expand=True)

        # 사용자 목록 표시
        tree = ttk.Treeview(list_frame, columns=("mac", "name", "role"), show="headings", height=15)
        tree.heading("mac", text="맥어드레스")
        tree.heading("name", text="사용자명")
        tree.heading("role", text="권한")
        
        # 컬럼 너비 설정
        tree.column("mac", width=200)
        tree.column("name", width=150)
        tree.column("role", width=100)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # 사용자 데이터 추가
        for user in users_data["users"]:
            tree.insert("", "end", values=(user["mac"], user["name"], user["role"]))
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 하단 닫기 버튼 영역 (항상 맨 아래)
        bottom_frame = tk.Frame(main_frame, bg='#f8f9fa')
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=12)
        close_btn = ModernButton(bottom_frame, text="닫기", command=dialog.destroy,
                                bg="#6B7280", fg="white", width=15, height=1)
        close_btn.pack()

    def add_user(self):
        """새 사용자 추가"""
        # 관리자 권한 확인
        if not self._check_admin_permission():
            return
        
        dialog = Toplevel(self.root)
        dialog.title("사용자 추가")
        dialog.geometry("450x400")  # 크기 증가
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#f8f9fa')
        
        # 중앙 정렬
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 150, self.root.winfo_rooty() + 150))
        
        # 메인 컨테이너
        main_frame = tk.Frame(dialog, bg='#f8f9fa')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 제목
        title_label = tk.Label(main_frame, text="새 사용자 추가", 
                              font=("맑은 고딕", 16, "bold"), bg='#f8f9fa', fg="#1f2937")
        title_label.pack(pady=(0, 25))
        
        # 입력 필드들
        input_frame = tk.Frame(main_frame, bg='#f8f9fa')
        input_frame.pack(fill=tk.X, pady=10)
        
        # 맥어드레스 입력
        tk.Label(input_frame, text="맥어드레스:", font=("맑은 고딕", 11, "bold"), 
                bg='#f8f9fa', fg="#374151").pack(anchor=tk.W)
        mac_entry = tk.Entry(input_frame, font=("맑은 고딕", 11), width=35)
        mac_entry.pack(fill=tk.X, pady=(5, 20))
        
        # 사용자명 입력
        tk.Label(input_frame, text="사용자명:", font=("맑은 고딕", 11, "bold"), 
                bg='#f8f9fa', fg="#374151").pack(anchor=tk.W)
        name_entry = tk.Entry(input_frame, font=("맑은 고딕", 11), width=35)
        name_entry.pack(fill=tk.X, pady=(5, 20))
        
        # 권한 선택 (더 명확하게 표시)
        tk.Label(input_frame, text="권한:", font=("맑은 고딕", 11, "bold"), 
                bg='#f8f9fa', fg="#374151").pack(anchor=tk.W, pady=(0, 10))
        
        role_var = tk.StringVar(value="user")
        role_frame = tk.Frame(input_frame, bg='#f8f9fa')
        role_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 라디오 버튼을 더 명확하게 표시
        user_radio = tk.Radiobutton(role_frame, text="일반 사용자", variable=role_var, value="user", 
                                   bg='#f8f9fa', fg="#374151", font=("맑은 고딕", 11),
                                   selectcolor="#E5E7EB", activebackground="#f8f9fa")
        user_radio.pack(side=tk.LEFT, padx=(0, 30))
        
        admin_radio = tk.Radiobutton(role_frame, text="관리자", variable=role_var, value="admin", 
                                    bg='#f8f9fa', fg="#374151", font=("맑은 고딕", 11),
                                    selectcolor="#E5E7EB", activebackground="#f8f9fa")
        admin_radio.pack(side=tk.LEFT)
        
        # 기본값 선택 표시
        user_radio.select()
        
        # 버튼 프레임 (명확하게 표시)
        button_frame = tk.Frame(main_frame, bg='#f8f9fa')
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))
        
        def save_user():
            mac = mac_entry.get().strip().upper()
            name = name_entry.get().strip()
            role = role_var.get()
            
            if not mac or not name:
                messagebox.showwarning("경고", "모든 필드를 입력해주세요.")
                return
            
            # 맥어드레스 형식 검증
            if not re.match(r'^([0-9A-F]{2}-){5}[0-9A-F]{2}$', mac):
                messagebox.showwarning("경고", "올바른 맥어드레스 형식을 입력해주세요.\n예: 80-E8-2C-EF-97-E0")
                return
            
            # 사용자 추가
            if self._add_user_to_file(mac, name, role):
                messagebox.showinfo("성공", f"사용자 '{name}'이(가) 추가되었습니다.")
                dialog.destroy()
            else:
                messagebox.showerror("오류", "사용자 추가에 실패했습니다.")
        
        # 저장/취소 버튼 (더 명확하게 표시)
        save_btn = ModernButton(button_frame, text="저장", command=save_user,
                               bg="#059669", fg="white", width=15, height=2,
                               font=("맑은 고딕", 11, "bold"))
        save_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        cancel_btn = ModernButton(button_frame, text="취소", command=dialog.destroy,
                                 bg="#6B7280", fg="white", width=15, height=2,
                                 font=("맑은 고딕", 11, "bold"))
        cancel_btn.pack(side=tk.LEFT)
        
        # 버튼 프레임에 테두리 추가 (시각적 구분)
        button_frame.configure(relief="solid", bd=1)

    def remove_user(self):
        """사용자 제거"""
        # 관리자 권한 확인
        if not self._check_admin_permission():
            return
        
        users_data = load_encrypted_users()
        if not users_data:
            messagebox.showwarning("경고", "사용자 정보를 불러올 수 없습니다.")
            return
        
        dialog = Toplevel(self.root)
        dialog.title("사용자 제거")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#f8f9fa')
        
        # 중앙 정렬
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 150, self.root.winfo_rooty() + 150))
        
        # 메인 컨테이너
        main_frame = tk.Frame(dialog, bg='#f8f9fa')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 제목
        title_label = tk.Label(main_frame, text="사용자 제거", 
                              font=("맑은 고딕", 16, "bold"), bg='#f8f9fa', fg="#1f2937")
        title_label.pack(pady=(0, 20))
        
        # 사용자 목록
        list_frame = tk.Frame(main_frame, bg='#f8f9fa')
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 리스트박스
        listbox = tk.Listbox(list_frame, font=("맑은 고딕", 10), height=12)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        
        # 사용자 데이터 추가
        for user in users_data["users"]:
            listbox.insert(tk.END, f"{user['name']} ({user['mac']}) - {user['role']}")
        
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 버튼 프레임
        button_frame = tk.Frame(main_frame, bg='#f8f9fa')
        button_frame.pack(pady=20)
        
        def remove_selected():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("경고", "제거할 사용자를 선택해주세요.")
                return
            
            index = selection[0]
            user = users_data["users"][index]
            
            if messagebox.askyesno("확인", f"사용자 '{user['name']}'을(를) 정말 제거하시겠습니까?"):
                if self._remove_user_from_file(index):
                    messagebox.showinfo("성공", f"사용자 '{user['name']}'이(가) 제거되었습니다.")
                    dialog.destroy()
                else:
                    messagebox.showerror("오류", "사용자 제거에 실패했습니다.")
        
        # 제거/취소 버튼
        remove_btn = ModernButton(button_frame, text="제거", command=remove_selected,
                                 bg="#DC2626", fg="white", width=12, height=1)
        remove_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        cancel_btn = ModernButton(button_frame, text="취소", command=dialog.destroy,
                                 bg="#6B7280", fg="white", width=12, height=1)
        cancel_btn.pack(side=tk.LEFT)

    def edit_users_json(self):
        """JSON 파일 편집"""
        # 관리자 권한 확인
        if not self._check_admin_permission():
            return
        
        users_data = load_encrypted_users()
        if not users_data:
            messagebox.showwarning("경고", "사용자 정보를 불러올 수 없습니다.")
            return
        
        dialog = Toplevel(self.root)
        dialog.title("JSON 파일 편집")
        dialog.geometry("700x600")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#f8f9fa')
        
        # 중앙 정렬
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 100))
        
        # 메인 컨테이너
        main_frame = tk.Frame(dialog, bg='#f8f9fa')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 제목
        title_label = tk.Label(main_frame, text="JSON 파일 편집", 
                              font=("맑은 고딕", 16, "bold"), bg='#f8f9fa', fg="#1f2937")
        title_label.pack(pady=(0, 15))
        
        # 안내 메시지
        info_label = tk.Label(main_frame, 
                             text="사용자 정보를 JSON 형식으로 편집할 수 있습니다.\n편집 후 '저장' 버튼을 클릭하세요.",
                             font=("맑은 고딕", 10), fg="#6B7280", bg='#f8f9fa')
        info_label.pack(pady=(0, 20))
        
        # JSON 편집 영역
        text_frame = tk.Frame(main_frame, bg='#f8f9fa')
        text_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 텍스트 위젯
        text_widget = tk.Text(text_frame, font=("Consolas", 10), wrap=tk.NONE)
        scrollbar_y = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar_x = tk.Scrollbar(text_frame, orient=tk.HORIZONTAL, command=text_widget.xview)
        text_widget.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # JSON 내용 표시
        json_content = json.dumps(users_data, ensure_ascii=False, indent=2)
        text_widget.insert(tk.END, json_content)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 버튼 프레임
        button_frame = tk.Frame(main_frame, bg='#f8f9fa')
        button_frame.pack(pady=20)
        
        def save_json():
            try:
                # JSON 파싱 테스트
                new_content = text_widget.get("1.0", tk.END).strip()
                new_data = json.loads(new_content)
                
                # 사용자 데이터 검증
                if "users" not in new_data or not isinstance(new_data["users"], list):
                    messagebox.showerror("오류", "올바른 사용자 데이터 형식이 아닙니다.")
                    return
                
                # 암호화하여 저장
                if save_encrypted_users(new_content):
                    messagebox.showinfo("성공", "JSON 파일이 저장되었습니다.")
                    dialog.destroy()
                else:
                    messagebox.showerror("오류", "파일 저장에 실패했습니다.")
                    
            except json.JSONDecodeError as e:
                messagebox.showerror("오류", f"JSON 형식이 올바르지 않습니다:\n{str(e)}")
            except Exception as e:
                messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{str(e)}")
        
        # 저장/취소 버튼
        save_btn = ModernButton(button_frame, text="저장", command=save_json,
                               bg="#059669", fg="white", width=12, height=1)
        save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        cancel_btn = ModernButton(button_frame, text="취소", command=dialog.destroy,
                                 bg="#6B7280", fg="white", width=12, height=1)
        cancel_btn.pack(side=tk.LEFT)

    def backup_users(self):
        """현재 사용자 정보를 백업 파일로 저장"""
        # 관리자 권한 확인
        if not self._check_admin_permission():
            return
        
        try:
            # 현재 사용자 정보 로드
            users_data = load_encrypted_users()
            if not users_data:
                messagebox.showerror("오류", "사용자 정보를 불러올 수 없습니다.")
                return
            
            # 백업 파일명 생성 (날짜 포함)
            from datetime import datetime
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"users_backup_{current_time}.enc"
            
            # 백업 파일 저장 위치 선택
            backup_path = filedialog.asksaveasfilename(
                title="백업 파일 저장 위치 선택",
                defaultextension=".enc",
                filetypes=[("Encrypted Files", "*.enc"), ("All Files", "*.*")],
                initialname=backup_filename
            )
            
            if not backup_path:
                return
            
            # 현재 암호화된 파일 존재 여부 확인
            if not os.path.exists("users.json.enc"):
                messagebox.showerror("오류", "users.json.enc 파일을 찾을 수 없습니다.\n먼저 사용자를 추가하거나 기존 파일을 확인해주세요.")
                return
            
            # 현재 암호화된 파일을 백업 위치로 복사
            import shutil
            try:
                shutil.copy2("users.json.enc", backup_path)
                print(f"백업 파일 복사 성공: {backup_path}")
            except Exception as copy_error:
                messagebox.showerror("오류", f"백업 파일 복사에 실패했습니다:\n{str(copy_error)}")
                print(f"파일 복사 오류: {copy_error}")
                return
            
            # 백업 정보 표시
            backup_info = f"""
백업이 완료되었습니다!

📁 백업 파일: {os.path.basename(backup_path)}
📍 저장 위치: {os.path.dirname(backup_path)}
📊 사용자 수: {len(users_data.get('users', []))}명
🕐 백업 시간: {current_time}

백업 파일을 안전한 곳에 보관하세요.
            """
            
            messagebox.showinfo("백업 완료", backup_info)
            
        except Exception as e:
            messagebox.showerror("오류", f"백업 중 오류가 발생했습니다:\n{str(e)}")
            print(f"백업 오류: {e}")
            # 상세한 오류 정보 출력
            import traceback
            traceback.print_exc()

    def restore_users_backup(self):
        """암호화된 백업 파일 복원"""
        # 관리자 권한 확인
        if not self._check_admin_permission():
            return
        
        file_path = filedialog.askopenfilename(
            title="복원할 암호화 파일 선택",
            filetypes=[("Encrypted Files", "*.enc"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # 백업 파일 정보 확인
            backup_users_data = load_encrypted_users(file_path)
            if not backup_users_data:
                messagebox.showerror("오류", "백업 파일을 읽을 수 없습니다.")
                return
            
            # 백업 파일 정보 표시
            backup_info = f"""
백업 파일 정보:

📁 파일명: {os.path.basename(file_path)}
📊 사용자 수: {len(backup_users_data.get('users', []))}명
📅 마지막 업데이트: {backup_users_data.get('last_updated', '알 수 없음')}
🔄 버전: {backup_users_data.get('version', '알 수 없음')}

사용자 목록:
"""
            for user in backup_users_data.get('users', []):
                backup_info += f"• {user['name']} ({user['mac']}) - {user['role']}\n"
            
            backup_info += "\n이 백업 파일로 복원하시겠습니까?"
            
            # 복원 확인
            if messagebox.askyesno("백업 복원 확인", backup_info):
                # 현재 파일 백업 (안전장치)
                from datetime import datetime
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                safety_backup = f"users_safety_backup_{current_time}.enc"
                import shutil
                shutil.copy2("users.json.enc", safety_backup)
                
                # 백업 파일 복원
                shutil.copy2(file_path, "users.json.enc")
                
                # 복원 완료 메시지
                restore_info = f"""
복원이 완료되었습니다!

✅ 백업 파일에서 복원 완료
📁 복원된 파일: {os.path.basename(file_path)}
🔄 안전 백업: {safety_backup}

프로그램을 재시작하면 변경사항이 적용됩니다.
                """
                
                messagebox.showinfo("복원 완료", restore_info)
                
        except Exception as e:
            messagebox.showerror("오류", f"복원 중 오류가 발생했습니다:\n{str(e)}")
            print(f"복원 오류: {e}")

    def _add_user_to_file(self, mac, name, role):
        """파일에 새 사용자 추가"""
        try:
            users_data = load_encrypted_users()
            if not users_data:
                return False
            
            # 중복 확인
            for user in users_data["users"]:
                if user["mac"] == mac:
                    messagebox.showwarning("경고", "이미 등록된 맥어드레스입니다.")
                    return False
            
            # 새 사용자 추가
            new_user = {"mac": mac, "name": name, "role": role}
            users_data["users"].append(new_user)
            users_data["last_updated"] = "2025-01-27"  # 현재 날짜로 업데이트
            
            # JSON 형식으로 변환하여 암호화 저장
            json_content = json.dumps(users_data, ensure_ascii=False, indent=2)
            return save_encrypted_users(json_content)
            
        except Exception as e:
            print(f"사용자 추가 실패: {e}")
            return False

    def _remove_user_from_file(self, index):
        """파일에서 사용자 제거"""
        try:
            users_data = load_encrypted_users()
            if not users_data:
                return False
            
            # 사용자 제거
            users_data["users"].pop(index)
            users_data["last_updated"] = "2025-01-27"  # 현재 날짜로 업데이트
            
            # JSON 형식으로 변환하여 암호화 저장
            json_content = json.dumps(users_data, ensure_ascii=False, indent=2)
            return save_encrypted_users(json_content)
            
        except Exception as e:
            print(f"사용자 제거 실패: {e}")
            return False

    def _check_admin_permission(self):
        """관리자 권한 확인"""
        try:
            # 현재 사용자의 맥어드레스 가져오기
            current_mac = get_mac_address()
            if not current_mac:
                messagebox.showerror("오류", "맥어드레스를 확인할 수 없습니다.")
                return False
            
            # 사용자 데이터에서 현재 사용자 정보 찾기
            users_data = load_encrypted_users()
            if not users_data:
                messagebox.showerror("오류", "사용자 정보를 불러올 수 없습니다.")
                return False
            
            # 현재 사용자 찾기
            current_user = None
            for user in users_data["users"]:
                if user["mac"] == current_mac:
                    current_user = user
                    break
            
            if not current_user:
                messagebox.showerror("오류", "등록되지 않은 사용자입니다.")
                return False
            
            # 관리자 권한 확인
            if current_user["role"] != "admin":
                messagebox.showwarning("권한 없음", "이 기능은 관리자만 사용할 수 있습니다.")
                return False
            
            return True
            
        except Exception as e:
            print(f"권한 확인 실패: {e}")
            messagebox.showerror("오류", "권한 확인 중 오류가 발생했습니다.")
            return False

    def show_help(self):
        """사용법 안내 (새로 추가)"""
        help_text = f"""
📖 Kunhwa PDF Editor {VERSION} 사용법

🚀 성능 설정
• 고성능 모드: 대용량 PDF 처리에 최적화 (권장)
• 균형 모드: 일반적인 용도에 적합
• 고품질 모드: 고해상도 썸네일

⌨️ 키보드 단축키
• Ctrl+O: PDF 열기
• Ctrl+S: PDF 저장
• Ctrl+클릭: 다중선택
• Shift+클릭: 범위선택
• Delete: 선택된 페이지 삭제
• Ctrl+A: 전체선택
• ESC: 선택해제
• Ctrl+휠(썸네일): 크기조정

📁 파일 작업
• PDF 열기/저장/병합
• 선택 페이지 저장
• 빈페이지 삽입
• 텍스트 추출

💡 성능 팁
• 대용량 PDF는 고성능 모드 사용
• 썸네일 크기는 Ctrl+휠로 조정 가능
• 메모리 부족 시 고성능 모드로 전환
        """
        messagebox.showinfo("사용법", help_text)

    def show_about(self):
        """프로그램 정보 (새로 추가)"""
        about_text = f"""
🎯 Kunhwa PDF Editor {VERSION}

📝 최근 업데이트: 2026-02-13 (v3.3)
🏢 개발: (주)건화 IT팀

✨ v3.3 신규 업데이트
• Undo/Redo 시스템 도입 (Ctrl+Z / Ctrl+Y, 최대 10단계)
• 페이지 이동 기능 추가 (Ctrl+G, Go To Page)
• 하단 상태표시줄 추가 (페이지 수, 파일 용량 표시)
• 최근 열었던 파일 목록 (파일 메뉴에서 확인 가능)
• 대용량 파일 처리 시 진행률 표시 (Progress Indicator)
• 사용자 인증 데이터 캐싱을 통한 메뉴 반응 속도 개선

🚀 핵심 최적화
• 메모리 캐시 자동 관리 및 최적화 엔진
• 드래그 앤 드롭 병합 및 교차 창 페이지 이동 강화
• 썸네일 생성 및 미리보기 선명도 자동 조정

Copyright 2026 Kunhwa Engineering & Consulting. All rights reserved.
        """
        messagebox.showinfo("프로그램 정보", about_text.strip())

    def copy_selected_pages(self, event=None):
        """선택된 페이지를 안전복사용 임시 PDF 바이트로 클립보드에 보관"""
        try:
            if not self.doc or not self.selected_indices:
                return
            indices = sorted(self.selected_indices)
            temp = fitz.open()
            # 연속 구간으로 묶어 삽입 최적화
            start = prev = indices[0]
            for idx in indices[1:] + [None]:
                if idx is None or idx != prev + 1:
                    temp.insert_pdf(self.doc, from_page=start, to_page=prev)
                    if idx is not None:
                        start = idx
                prev = idx if idx is not None else prev
            # 바이트 및 임시 파일로 저장 (교차 프로세스/창 호환)
            self.page_clipboard_bytes = temp.write()
            try:
                fd, tmp_path = tempfile.mkstemp(prefix="kunhwa_pdf_clip_", suffix=".pdf")
                os.close(fd)
                temp.save(tmp_path)
                # OS 클립보드에는 경로 텍스트로 저장
                self.root.clipboard_clear()
                self.root.clipboard_append(tmp_path)
                print(f"클립보드 파일 경로: {tmp_path}")
            except Exception as e:
                print(f"임시 파일 저장/클립보드 경로 저장 실패: {e}")
            temp.close()
            print(f"클립보드에 {len(indices)}개 페이지 저장")
        except Exception as e:
            print(f"페이지 복사 실패: {e}")

    def paste_pages_from_clipboard(self, event=None):
        """클립보드의 임시 PDF 바이트를 현재 문서에 붙여넣기"""
        if self.doc:
            self.undo_manager.save_state(self.doc)
        try:
            # 1) 우선 OS 클립보드의 경로 시도
            path = None
            try:
                path = self.root.clipboard_get()
            except Exception:
                path = None
            # 2) 경로가 유효하지 않으면 메모리 바이트 사용
            if path and os.path.exists(path):
                src = fitz.open(path)
            elif self.page_clipboard_bytes:
                src = fitz.open(stream=self.page_clipboard_bytes, filetype='pdf')
            else:
                return
            if not self.doc:
                self.doc = fitz.open()
                self.current_page_index = 0
            # 붙여넣기 위치: 현재 선택이 있으면 그 앞, 없으면 마지막
            if self.selected_indices:
                insert_pos = min(self.selected_indices)
            else:
                insert_pos = len(self.doc)
            # 전체를 대상에 삽입
            self.doc.insert_pdf(src, from_page=0, to_page=len(src)-1, start_at=insert_pos)
            src.close()
            self.refresh_thumbnails()
            self.update_preview()
            print("클립보드에서 페이지 붙여넣기 완료")
        except Exception as e:
            print(f"페이지 붙여넣기 실패: {e}")

    # ═══════════════════════════════════════════════════
    # v3.3 신규 메서드
    # ═══════════════════════════════════════════════════

    def perform_undo(self):
        """Undo: 이전 상태로 되돌리기 (Ctrl+Z)"""
        if not self.doc:
            return
        restored = self.undo_manager.undo(self.doc)
        if restored:
            self.doc = restored
            self.current_page_index = min(self.current_page_index, len(self.doc) - 1)
            self.selected_indices.clear()
            self._thumbnail_cache.clear()
            self._preview_cache.clear()
            self.refresh_thumbnails()
            self.update_preview()
            self._update_status_bar()
            print("Undo 수행 완료")
        else:
            print("더 이상 되돌릴 수 없습니다.")

    def perform_redo(self):
        """Redo: 다시 실행 (Ctrl+Y)"""
        if not self.doc:
            return
        restored = self.undo_manager.redo(self.doc)
        if restored:
            self.doc = restored
            self.current_page_index = min(self.current_page_index, len(self.doc) - 1)
            self.selected_indices.clear()
            self._thumbnail_cache.clear()
            self._preview_cache.clear()
            self.refresh_thumbnails()
            self.update_preview()
            self._update_status_bar()
            print("Redo 수행 완료")
        else:
            print("더 이상 다시 실행할 수 없습니다.")

    def show_goto_page_dialog(self):
        """Go To Page 다이얼로그 (Ctrl+G)"""
        if not self.doc:
            messagebox.showwarning("경고", "PDF를 먼저 열어주세요.")
            return
        
        dialog = Toplevel(self.root)
        dialog.title("페이지 이동")
        dialog.geometry("340x180")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#f8f9fa')
        dialog.resizable(False, False)
        
        # 중앙 정렬
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 340) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 180) // 2
        dialog.geometry(f"+{x}+{y}")
        
        total = len(self.doc)
        
        tk.Label(dialog, text="페이지 이동", font=("맑은 고딕", 14, "bold"),
                bg='#f8f9fa', fg='#212529').pack(pady=(15, 5))
        
        tk.Label(dialog, text=f"이동할 페이지 번호 (1 ~ {total}):",
                font=("맑은 고딕", 10), bg='#f8f9fa', fg='#495057').pack(pady=5)
        
        entry = tk.Entry(dialog, font=("맑은 고딕", 12), width=10, justify='center')
        entry.pack(pady=5)
        entry.focus_set()
        
        def go_to_page():
            try:
                page_num = int(entry.get())
                if 1 <= page_num <= total:
                    self.current_page_index = page_num - 1
                    self.selected_indices = {page_num - 1}
                    self.update_preview()
                    self.refresh_thumbnails()
                    self._update_status_bar()
                    dialog.destroy()
                else:
                    messagebox.showwarning("경고", f"1~{total} 범위의 숫자를 입력하세요.", parent=dialog)
            except ValueError:
                messagebox.showwarning("경고", "올바른 숫자를 입력하세요.", parent=dialog)
        
        btn_frame = tk.Frame(dialog, bg='#f8f9fa')
        btn_frame.pack(pady=10)
        
        ModernButton(btn_frame, text="이동", command=go_to_page,
                    bg="#2563EB", fg="white").pack(side=tk.LEFT, padx=5)
        ModernButton(btn_frame, text="취소", command=dialog.destroy,
                    bg="#6B7280", fg="white").pack(side=tk.LEFT, padx=5)
        
        entry.bind("<Return>", lambda e: go_to_page())

    def _update_status_bar(self):
        """하단 상태표시줄 업데이트"""
        try:
            if self.doc:
                self._status_pages.config(text=f"페이지: {len(self.doc)}")
                # 파일 크기 계산
                try:
                    pdf_bytes = self.doc.tobytes()
                    size_kb = len(pdf_bytes) / 1024
                    if size_kb > 1024:
                        self._status_filesize.config(text=f"파일크기: {size_kb/1024:.1f}MB")
                    else:
                        self._status_filesize.config(text=f"파일크기: {size_kb:.0f}KB")
                except Exception:
                    self._status_filesize.config(text="파일크기: -")
            else:
                self._status_pages.config(text="페이지: 0")
                self._status_filesize.config(text="파일크기: -")
            
            # 줌 레벨
            zoom_pct = int(self.preview_scale * 100)
            self._status_zoom.config(text=f"확대: {zoom_pct}%")
            
            # Undo 정보
            undo_count = len(self.undo_manager._undo_stack)
            redo_count = len(self.undo_manager._redo_stack)
            if undo_count > 0 or redo_count > 0:
                self._status_undo.config(text=f"↩ Undo: {undo_count} | Redo: {redo_count} ↪")
            else:
                self._status_undo.config(text="")
        except Exception as e:
            print(f"상태표시줄 업데이트 오류: {e}")

    def _refresh_recent_files_menu(self):
        """최근 파일 메뉴 갱신"""
        try:
            if not hasattr(self, '_recent_menu'):
                return
            self._recent_menu.delete(0, tk.END)
            recent = self.recent_files_manager.files
            if not recent:
                self._recent_menu.add_command(label="(없음)", state="disabled")
            else:
                for fp in recent:
                    label = os.path.basename(fp)
                    self._recent_menu.add_command(
                        label=label,
                        command=lambda p=fp: self._open_recent_file(p)
                    )
        except Exception as e:
            print(f"최근 파일 메뉴 갱신 오류: {e}")

    def _open_recent_file(self, file_path):
        """최근 파일 열기"""
        if os.path.exists(file_path):
            self.open_pdf_from_path(file_path)
        else:
            messagebox.showwarning("경고", f"파일을 찾을 수 없습니다:\n{file_path}")
            self.recent_files_manager.remove(file_path)

    def _on_close_window(self):
        """창 종료 시 전역 레지스트리에서 제거"""
        try:
            if self in OPEN_EDITORS:
                OPEN_EDITORS.remove(self)
        except Exception:
            pass
        self.root.destroy()

    def new_window(self):
        """빈 새 창 열기 (exe 패킹 후에도 동작)"""
        try:
            subprocess.Popen(_build_launch_command())
        except Exception as e:
            messagebox.showerror("오류", f"새 창 실행에 실패했습니다.\n{e}")

    def new_window_with_file(self):
        """파일 선택 후 새 창으로 열기"""
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not path:
            return
        try:
            subprocess.Popen(_build_launch_command([path]))
        except Exception as e:
            messagebox.showerror("오류", f"새 창 실행에 실패했습니다.\n{e}")

    def new_window_in_process(self):
        """같은 프로세스 내에 Toplevel로 새 편집 창 생성 (DnD 호환성 높음)"""
        try:
            # 새 최상위 창
            top = tk.Toplevel(self.root)
            # 같은 프로세스 내에서 새 에디터 인스턴스
            PDFEditorApp(top)
            print("같은 프로세스 새 창 생성 완료")
        except Exception as e:
            print(f"같은 프로세스 새 창 생성 실패: {e}")

if __name__ == "__main__":
    # 사용자 인증 확인
    if not check_authorization():
        print("인증 실패로 프로그램을 종료합니다.")
        sys.exit(1)
    
    # 인증 성공 시 프로그램 실행
    # 멀티 창 실행: 파일 여러 개 동시에 열기 지원 (빈 실행도 가능)
    def launch_new_editor(initial_path: str | None = None):
        r = TkinterDnD.Tk() if DRAG_DROP_AVAILABLE else tk.Tk()
        r.geometry("1300x800")
        app = PDFEditorApp(r)
        if initial_path and os.path.exists(initial_path):
            try:
                app.open_pdf_from_path(initial_path)
            except Exception:
                pass
        r.mainloop()

    # 커맨드라인 인자로 넘어온 첫 번째 파일을 현재 프로세스에서 직접 열기
    initial_file = sys.argv[1] if len(sys.argv) > 1 else None
    if initial_file and os.path.exists(initial_file):
        launch_new_editor(initial_file)
    else:
        launch_new_editor()