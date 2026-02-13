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
REMOTE_AUTH_URL = "https://raw.githubusercontent.com/TaeW81/PDF_editor/main/users.json.enc"

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

def encrypt_data(text, password=PASSWORD):
    """Encrypts text using simple XOR + Base64."""
    try:
        # Ensure text is string before encoding
        if not isinstance(text, str):
            text = json.dumps(text)
            
        data_bytes = text.encode('utf-8')
        encrypted = bytearray()
        for i, byte in enumerate(data_bytes):
            key_byte = ord(password[i % len(password)])
            encrypted.append(byte ^ key_byte)
        return base64.b64encode(encrypted).decode('utf-8')
    except Exception as e:
        print(f"Encryption failed: {e}")
        return None

def decrypt_data(encrypted_text, password=PASSWORD):
    """Decrypts text."""
    try:
        encrypted = base64.b64decode(encrypted_text.encode('utf-8'))
        decrypted = bytearray()
        for i, byte in enumerate(encrypted):
            key_byte = ord(password[i % len(password)])
            decrypted.append(byte ^ key_byte)
        return decrypted.decode('utf-8')
    except Exception as e:
        print(f"Decryption failed: {e}")
        return None

User = collections.namedtuple('User', ['name', 'role', 'mac'])

class AuthManager:
    def __init__(self, users_file=USERS_FILE):
        self.users_file = users_file
        self.current_user = None

    def load_users(self):
        """Loads users from encrypted file."""
        if not os.path.exists(self.users_file):
            return None
        
        try:
            with open(self.users_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            decrypted = decrypt_data(content)
            if decrypted:
                return json.loads(decrypted)
        except Exception as e:
            print(f"Error loading users: {e}")
            return None

    def fetch_remote_users(self):
        """Fetches users file from remote server."""
        if not REMOTE_AUTH_URL:
            return False
            
        try:
            with urllib.request.urlopen(REMOTE_AUTH_URL, timeout=5) as response:
                if response.status == 200:
                    data = response.read().decode('utf-8')
                    # Validate if it decodes?
                    # For now, just save it to local file
                    with open(self.users_file, 'w', encoding='utf-8') as f:
                        f.write(data)
                    return True
        except Exception as e:
            print(f"Remote auth failed: {e}")
            return False

    def authenticate(self):
        """Checks if current machine is authorized."""
        # Try to sync with server first
        if REMOTE_AUTH_URL:
            print(f"Checking remote auth: {REMOTE_AUTH_URL}")
            self.fetch_remote_users()
            
        current_mac = get_mac_address()
        if not current_mac:
            return False, "Could not determine MAC address."

        users_data = self.load_users()
        if not users_data or "users" not in users_data:
            # Check if file serves purely as a key or requires specific entries
            # For this app, we need the file to exist and contain the user
            return False, f"User database not found at {self.users_file}."

        for user in users_data["users"]:
            if user["mac"] == current_mac:
                self.current_user = User(name=user["name"], role=user["role"], mac=user["mac"])
                return True, f"Welcome, {user['name']}"

        return False, f"Access Denied. MAC: {current_mac}"

    def get_current_user_name(self):
        return self.current_user.name if self.current_user else "Unknown"

    def is_admin(self):
        return self.current_user and self.current_user.role == 'admin'

    def save_users(self, users_data):
        """Saves users data to encrypted file."""
        try:
            content = json.dumps(users_data, ensure_ascii=False, indent=2)
            encrypted = encrypt_data(content)
            
            with open(self.users_file, 'w', encoding='utf-8') as f:
                f.write(encrypted)
            return True
        except Exception as e:
            print(f"Error saving users: {e}")
            return False

    def add_user(self, name, role, mac):
        """Adds a new user."""
        users_data = self.load_users()
        if not users_data:
            users_data = {"users": [], "version": "1.0", "last_updated": ""}
            
        # Check duplicate
        for user in users_data["users"]:
            if user["mac"] == mac:
                return False, "MAC Address already exists."
        
        users_data["users"].append({
            "name": name,
            "role": role,
            "mac": mac
        })
        
        if self.save_users(users_data):
            return True, "User added successfully."
        return False, "Failed to save user data."

    def remove_user(self, mac):
        """Removes a user by MAC."""
        users_data = self.load_users()
        if not users_data:
            return False, "No user data found."
            
        initial_count = len(users_data["users"])
        users_data["users"] = [u for u in users_data["users"] if u["mac"] != mac]
        
        if len(users_data["users"]) == initial_count:
            return False, "User not found."
            
        if self.save_users(users_data):
            return True, "User removed successfully."
        return False, "Failed to save user data."

    def get_all_users(self):
        """Returns list of all users."""
        data = self.load_users()
        return data.get("users", []) if data else []
        
    def backup_users(self, backup_path):
        import shutil
        try:
            if os.path.exists(self.users_file):
                shutil.copy2(self.users_file, backup_path)
                return True, "Backup successful."
            return False, "User file does not exist."
        except Exception as e:
            return False, str(e)

    def restore_users(self, backup_path):
        import shutil
        try:
            if os.path.exists(backup_path):
                shutil.copy2(backup_path, self.users_file)
                return True, "Restore successful."
            return False, "Backup file not found."
        except Exception as e:
            return False, str(e)
