import sys
import os
import subprocess
from urllib.parse import unquote

DOPUS_RT = r"C:\Program Files\GPSoftware\Directory Opus\dopusrt.exe"
DOPUS_EXE = r"C:\Program Files\GPSoftware\Directory Opus\dopus.exe"

def main():
    if len(sys.argv) < 2:
        return
        
    raw_url = sys.argv[1]
    # Example raw_url: openopus://D:/Active%20research/OpenSearch/indexer_config.json
    if "://" in raw_url:
        protocol, path_part = raw_url.split("://", 1)
    else:
        protocol, path_part = "openfile", raw_url
        
    file_path = unquote(path_part).strip('/')
    # On Windows, fix drive letter path e.g. D:/Active research -> D:\Active research
    file_path = os.path.normpath(file_path)
    
    # Fix single letter drive formatting if needed e.g. D:\...
    if len(file_path) >= 2 and file_path[1] == ':' and not file_path.startswith('\\\\'):
        pass
    elif len(file_path) >= 2 and file_path[1] == '/':
        file_path = file_path[0] + ':' + file_path[2:]

    file_exists = os.path.exists(file_path)
    target_path = file_path

    if not file_exists:
        parent_dir = os.path.dirname(file_path)
        if os.path.exists(parent_dir):
            target_path = parent_dir
        else:
            return

    if protocol == "openfile":
        if file_exists:
            os.startfile(target_path)
        else:
            subprocess.Popen(['explorer', target_path])

    elif protocol == "openexplorer":
        if file_exists:
            subprocess.Popen(['explorer', '/select,', target_path])
        else:
            subprocess.Popen(['explorer', target_path])

    elif protocol == "openopus":
        if os.path.exists(DOPUS_RT):
            if file_exists and not os.path.isdir(target_path):
                parent_dir = os.path.dirname(target_path)
                file_name = os.path.basename(target_path)
                subprocess.Popen([DOPUS_RT, "/cmd", "Go", parent_dir, "NEW", f"SELECT={file_name}"])
            else:
                subprocess.Popen([DOPUS_RT, "/cmd", "Go", target_path, "NEW"])
        elif os.path.exists(DOPUS_EXE):
            if file_exists and not os.path.isdir(target_path):
                parent_dir = os.path.dirname(target_path)
                subprocess.Popen([DOPUS_EXE, parent_dir])
            else:
                subprocess.Popen([DOPUS_EXE, target_path])
        else:
            if file_exists and not os.path.isdir(target_path):
                subprocess.Popen(['explorer', '/select,', target_path])
            else:
                subprocess.Popen(['explorer', target_path])

if __name__ == "__main__":
    main()
