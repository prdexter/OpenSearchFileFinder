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

    if not os.path.exists(file_path):
        return

    if protocol == "openfile":
        os.startfile(file_path)

    elif protocol == "openexplorer":
        subprocess.Popen(['explorer', '/select,', file_path])

    elif protocol == "openopus":
        if os.path.exists(DOPUS_RT):
            subprocess.Popen([DOPUS_RT, "/cmd", "Go", file_path, "NEW", "SELECT"])
        elif os.path.exists(DOPUS_EXE):
            subprocess.Popen([DOPUS_EXE, "/select", file_path])
        else:
            subprocess.Popen(['explorer', '/select,', file_path])

if __name__ == "__main__":
    main()
