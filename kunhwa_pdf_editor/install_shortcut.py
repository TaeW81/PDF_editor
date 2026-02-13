"""Windows 바탕화면 및 시작메뉴에 바로가기 생성"""
import os
import sys
import shutil


def _get_icon_path():
    """패키지 내 아이콘 파일 경로 반환"""
    pkg_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(pkg_dir, "data", "kunhwa_logo.ico")
    if os.path.exists(icon_path):
        return icon_path
    return None


def _get_script_path():
    """kunhwa-pdf-editor 실행 스크립트 경로 반환"""
    # pip install 후 Scripts 폴더에 생성되는 실행 파일 경로
    scripts_dir = os.path.join(os.path.dirname(sys.executable), "Scripts")
    
    # .exe 확인 (Windows)
    exe_path = os.path.join(scripts_dir, "kunhwa-pdf-editor.exe")
    if os.path.exists(exe_path):
        return exe_path
    
    # gui-scripts 확인
    gui_exe_path = os.path.join(scripts_dir, "kunhwa-pdf-editor-gui.exe")
    if os.path.exists(gui_exe_path):
        return gui_exe_path
    
    # 스크립트 파일 확인
    script_path = os.path.join(scripts_dir, "kunhwa-pdf-editor")
    if os.path.exists(script_path):
        return script_path
    
    return None


def _create_shortcut_with_win32(shortcut_path, target_path, icon_path=None):
    """win32com을 사용하여 .lnk 바로가기 생성"""
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = target_path
        shortcut.WorkingDirectory = os.path.expanduser("~")
        shortcut.Description = "Kunhwa PDF Editor v3.2"
        if icon_path and os.path.exists(icon_path):
            shortcut.IconLocation = icon_path
        shortcut.save()
        return True
    except ImportError:
        return False


def _create_shortcut_with_powershell(shortcut_path, target_path, icon_path=None):
    """PowerShell을 사용하여 .lnk 바로가기 생성 (pywin32 없을 때 대체)"""
    import subprocess
    
    icon_cmd = ""
    if icon_path and os.path.exists(icon_path):
        icon_cmd = f'$s.IconLocation = "{icon_path}"'
    
    ps_script = f'''
$WshShell = New-Object -comObject WScript.Shell
$s = $WshShell.CreateShortcut("{shortcut_path}")
$s.TargetPath = "{target_path}"
$s.WorkingDirectory = "{os.path.expanduser("~")}"
$s.Description = "Kunhwa PDF Editor v3.2"
{icon_cmd}
$s.Save()
'''
    
    try:
        result = subprocess.run(
            ["powershell", "-Command", ps_script],
            capture_output=True, text=True, encoding="utf-8"
        )
        return result.returncode == 0
    except Exception:
        return False


def create_shortcuts():
    """바탕화면과 시작메뉴에 바로가기를 생성합니다."""
    print("=" * 50)
    print("  Kunhwa PDF Editor - 바로가기 생성")
    print("=" * 50)
    
    # 실행 파일 경로 찾기
    target_path = _get_script_path()
    if not target_path:
        print("\n[X] 실행 파일을 찾을 수 없습니다.")
        print("    pip install 이 정상적으로 완료되었는지 확인하세요.")
        return False
    
    print(f"\n[*] 실행 파일: {target_path}")
    
    # 아이콘 경로
    icon_path = _get_icon_path()
    if icon_path:
        print(f"[*] 아이콘: {icon_path}")
    
    success_count = 0
    
    # 1. 바탕화면 바로가기
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.exists(desktop_dir):
        # 한국어 Windows의 경우
        desktop_dir = os.path.join(os.path.expanduser("~"), "바탕 화면")
    
    if os.path.exists(desktop_dir):
        desktop_shortcut = os.path.join(desktop_dir, "Kunhwa PDF Editor.lnk")
        print(f"\n[>] 바탕화면 바로가기 생성 중...")
        
        created = _create_shortcut_with_win32(desktop_shortcut, target_path, icon_path)
        if not created:
            created = _create_shortcut_with_powershell(desktop_shortcut, target_path, icon_path)
        
        if created:
            print(f"    [OK] 생성 완료: {desktop_shortcut}")
            success_count += 1
        else:
            print(f"    [X] 생성 실패")
    
    # 2. 시작메뉴 바로가기
    start_menu_dir = os.path.join(
        os.environ.get("APPDATA", ""),
        "Microsoft", "Windows", "Start Menu", "Programs"
    )
    
    if os.path.exists(start_menu_dir):
        start_shortcut = os.path.join(start_menu_dir, "Kunhwa PDF Editor.lnk")
        print(f"\n[>] 시작메뉴 바로가기 생성 중...")
        
        created = _create_shortcut_with_win32(start_shortcut, target_path, icon_path)
        if not created:
            created = _create_shortcut_with_powershell(start_shortcut, target_path, icon_path)
        
        if created:
            print(f"    [OK] 생성 완료: {start_shortcut}")
            success_count += 1
        else:
            print(f"    [X] 생성 실패")
    
    # 결과 요약
    print(f"\n{'=' * 50}")
    if success_count > 0:
        print(f"  [OK] {success_count}개의 바로가기가 생성되었습니다!")
        print(f"  바탕화면 또는 시작메뉴에서 실행할 수 있습니다.")
    else:
        print(f"  [X] 바로가기 생성에 실패했습니다.")
        print(f"  수동으로 실행: kunhwa-pdf-editor")
    print(f"{'=' * 50}")
    
    return success_count > 0


if __name__ == "__main__":
    create_shortcuts()
