import os
import sys
import winreg

def register_custom_protocol(protocol_name, handler_cmd):
    key_path = f"Software\\Classes\\{protocol_name}"
    try:
        # Create or open key under HKCU (no admin required!)
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f"URL:{protocol_name} Protocol")
            winreg.SetValueEx(key, "URL Protocol", 0, winreg.REG_SZ, "")
            
            with winreg.CreateKey(key, r"shell\open\command") as cmd_key:
                winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, handler_cmd)
        print(f"[+] Registered custom protocol: {protocol_name}://")
    except Exception as e:
        print(f"[-] Error registering {protocol_name}: {e}")

if __name__ == "__main__":
    py_exe = sys.executable
    handler_script = os.path.abspath(os.path.join(os.path.dirname(__file__), "protocol_handler.py"))
    
    cmd_str = f'"{py_exe}" "{handler_script}" "%1"'
    
    register_custom_protocol("openfile", cmd_str)
    register_custom_protocol("openopus", cmd_str)
    register_custom_protocol("openexplorer", cmd_str)
