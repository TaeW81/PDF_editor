import os
import sys
import json
import base64
import subprocess
import re
import uuid
import collections
import urllib.request
import urllib.error

# Import from config
# Need to add parent directory to path to import config if running as script
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config.settings import USERS_FILE

PASSWORD = "KUNHWA2025"
REMOTE_AUTH_URL = "https://gist.githubusercontent.com/TaeW81/8c6597546e977140599d675c4760c298/raw"

def get_mac_address():
    """Retrieves the system's MAC address with caching."""
    if hasattr(get_mac_address, '_cached_mac'):
        return get_mac_address._cached_mac
        
    try:
        # Try ipconfig method first (Windows)
        result = subprocess.run(['ipconfig', '/all'], capture_output=True, text=True, encoding='cp949')
        if result.returncode == 0:
            lines = result.stdout.split('\n')
            for line in lines:
                if '물리적 주소' in line or 'Physical Address' in line:
                    mac_match = re.search(r'([0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2})', line, re.IGNORECASE)
                    if mac_match:
                        mac = mac_match.group(1).upper()
                        get_mac_address._cached_mac = mac
                        return mac
                        
        # Fallback to uuid method
        mac = uuid.getnode()
        mac_address = ':'.join(['{:02x}'.format((mac >> elements) & 0xff) for elements in range(0,2*6,2)][::-1])
        mac = mac_address.replace(':', '-').upper()
        get_mac_address._cached_mac = mac
        return mac
        
    except Exception as e:
        print(f"Failed to get MAC address: {e}")
        return None



User = collections.namedtuple('User', ['name', 'role', 'mac'])


class AuthManager:
    def __init__(self, users_file=USERS_FILE):
        # We no longer use users_file for storage, but keep arg using it or just ignore
        self.users_file = users_file
        self.users_data = None # In-memory storage
        self.current_user = None
        
        # Cleanup local file if exists (Security)
        if os.path.exists(self.users_file):
            try:
                os.remove(self.users_file)
                print(f"Removed local user file: {self.users_file}")
            except Exception as e:
                print(f"Failed to remove local user file: {e}")

    def load_users(self):
        """Returns in-memory users data."""
        return self.users_data

    def fetch_remote_users(self):
        """Fetches users from remote server and stores in memory."""
        if not REMOTE_AUTH_URL:
            return False
            
        try:
            # Cache busting
            import time
            url_with_cache_bust = f"{REMOTE_AUTH_URL}?t={int(time.time())}"
            print(f"Fetching from: {url_with_cache_bust}")
            
            with urllib.request.urlopen(url_with_cache_bust, timeout=10) as response:
                if response.status == 200:
                    raw_data = response.read()
                    
                    # Robust Decoding
                    data = None
                    for encoding in ['utf-8', 'cp949', 'euc-kr']:
                        try:
                            data = raw_data.decode(encoding).strip()
                            break
                        except UnicodeDecodeError:
                            continue
                            
                    if data is None:
                        # Final fallback
                        data = raw_data.decode('utf-8', errors='ignore').strip()
                    
                    try:
                        # Direct JSON parse (No encryption)
                        self.users_data = json.loads(data)
                        print("Remote auth sync successful (Plain JSON).")
                        return True
                    except json.JSONDecodeError as e:
                        print(f"Remote data is not valid JSON: {e}")
                        # print(data[:100]) # Debug
                        return False
        except Exception as e:
            print(f"Remote auth failed: {e}")
            return False

    def authenticate(self):
        """Checks if current machine is authorized."""
        
        # Try to sync with server first
        # In In-Memory mode, if fetch fails and we have no data, we cannot auth.
        success = self.fetch_remote_users()
        
        if not success and not self.users_data:
             return False, "서버 연결 실패. 인터넷 연결을 확인해주세요."
            
        current_mac = get_mac_address()
        if not current_mac:
            return False, "MAC 주소를 확인할 수 없습니다."

        users_data = self.load_users()
        if not users_data or "users" not in users_data:
            return False, f"사용자 정보가 없습니다.\nMAC: {current_mac}"

        for user in users_data["users"]:
            if user["mac"] == current_mac:
                self.current_user = User(name=user["name"], role=user["role"], mac=user["mac"])
                return True, f"반갑습니다, {user['name']}님"

        return False, f"인증되지 않은 사용자입니다.\nMAC: {current_mac}"

    def get_current_user_name(self):
        return self.current_user.name if self.current_user else "Unknown"

    def is_admin(self):
        return self.current_user and self.current_user.role == 'admin'

    def save_users(self, users_data):
        """Updates in-memory data (Does not save to file)."""
        self.users_data = users_data
        return True

    def get_users_json_string(self):
        """Returns plain JSON string of current memory data for Export."""
        if not self.users_data:
            return ""
        try:
            return json.dumps(self.users_data, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"JSON dump failed: {e}")
            return ""

    def add_user(self, name, role, mac):
        """Adds a new user to memory."""
        users_data = self.load_users()
        if not users_data:
            users_data = {"users": [], "version": "1.0", "last_updated": ""}
            
        # Check duplicate
        for user in users_data["users"]:
            if user["mac"] == mac:
                return False, "이미 등록된 MAC 주소입니다."
        
        users_data["users"].append({
            "name": name,
            "role": role,
            "mac": mac
        })
        
        self.save_users(users_data)
        return True, "사용자가 추가되었습니다. (메모리 반영됨)"


    def update_user(self, original_mac, new_name, new_role, new_mac):
        """Updates an existing user."""
        users_data = self.load_users()
        if not users_data:
            return False, "사용자 데이터가 없습니다."
            
        # Find user
        target_user = None
        for u in users_data["users"]:
            if u["mac"] == original_mac:
                target_user = u
                break
                
        if not target_user:
            return False, "사용자를 찾을 수 없습니다."
            
        # Check duplicate MAC if changed
        if original_mac != new_mac:
            for u in users_data["users"]:
                if u["mac"] == new_mac:
                    return False, "이미 존재하는 MAC 주소입니다."
                    
        # Update
        target_user["name"] = new_name
        target_user["role"] = new_role
        target_user["mac"] = new_mac
        
        self.save_users(users_data)
        return True, "사용자 정보가 수정되었습니다."

    def remove_user(self, mac):
        """Removes a user by MAC from memory."""
        users_data = self.load_users()
        if not users_data:
            return False, "사용자 데이터가 없습니다."
            
        initial_count = len(users_data["users"])
        users_data["users"] = [u for u in users_data["users"] if u["mac"] != mac]
        
        if len(users_data["users"]) == initial_count:
            return False, "사용자를 찾을 수 없습니다."
            
        self.save_users(users_data)
        return True, "사용자가 삭제되었습니다. (메모리 반영됨)"


    def get_all_users(self):
        """Returns list of all users."""
        data = self.load_users()
        return data.get("users", []) if data else []
