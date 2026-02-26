import sys
import os

# Add project root to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.auth import AuthManager

def register_current_pc():
    auth = AuthManager()
    
    # Specific MAC from the screenshot
    target_mac = "30-56-0F-5D-7E-02" 
    name = "Developer"
    role = "admin"
    
    print(f"Adding user: {name} ({target_mac})")
    
    # Check if file exists, if not create empty
    if not os.path.exists(auth.users_file):
        print("Creating new users file...")
        auth.save_users({"users": [], "version": "1.0"})
        
    success, msg = auth.add_user(name, role, target_mac)
    print(f"Result: {msg}")

if __name__ == "__main__":
    register_current_pc()
