import os
import subprocess
import shutil

# ================= 配置区域 =================
PY_FILE = "PDF工具箱.py"
EXE_NAME = "PDF工具箱"
ISCC_PATH = r"D:\App\Innosetup\Inno Setup 6\ISCC.exe" # 确认路径
ISS_FILE = "PDF工具箱安装包.iss"
VERSION_FILE = "version.txt"
# ===========================================

def get_update_v():
    if not os.path.exists(VERSION_FILE):
        with open(VERSION_FILE, "w") as f: f.write("1.0.0"); return "1.0.0"
    with open(VERSION_FILE, "r") as f: v = f.read().strip()
    p = v.split('.'); p[-1] = str(int(p[-1]) + 1); nv = ".".join(p)
    with open(VERSION_FILE, "w") as f: f.write(nv); return nv

def main():
    ver = get_update_v()
    [shutil.rmtree(f) for f in ['dist', 'build'] if os.path.exists(f)]
    
    # 1. PyInstaller 打包
    cmd = f'pyinstaller --noconsole --onefile --clean --collect-all customtkinter --collect-all googletrans --name "{EXE_NAME}" "{PY_FILE}"'
    if not subprocess.run(cmd, shell=True).returncode == 0: return

    # 2. 生成 ISS 脚本 (可选路径与快捷方式)
    exe_abs = os.path.abspath(f"dist\\{EXE_NAME}.exe")
    iss_content = f"""
[Setup]
AppName={EXE_NAME}
AppVersion={ver}
DefaultDirName={{autopf}}\\{EXE_NAME}
DisableDirPage=no
OutputBaseFilename={EXE_NAME}_Setup_v{ver}
Compression=lzma
SolidCompression=yes
OutputDir={os.getcwd()}

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "{{cm:CreateDesktopIcon}}"; GroupDescription: "{{cm:AdditionalIcons}}"; Flags: unchecked

[Files]
Source: "{exe_abs}"; DestDir: "{{app}}"; Flags: ignoreversion

[Icons]
Name: "{{group}}\\{EXE_NAME}"; Filename: "{{app}}\\{EXE_NAME}.exe"
Name: "{{autodesktop}}\\{EXE_NAME}"; Filename: "{{app}}\\{EXE_NAME}.exe"; Tasks: desktopicon

[Registry]
Root: HKCR; Subkey: "SystemFileAssociations\\.docx\\shell\\PDFToolBox"; ValueType: string; ValueData: "PDF工具箱处理"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\\.docx\\shell\\PDFToolBox\\command"; ValueType: string; ValueData: \"""{{app}}\\{EXE_NAME}.exe"" ""%1\"""; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\\.pdf\\shell\\PDFToolBox"; ValueType: string; ValueData: "PDF工具箱处理"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\\.pdf\\shell\\PDFToolBox\\command"; ValueType: string; ValueData: \"""{{app}}\\{EXE_NAME}.exe"" ""%1\"""; Flags: uninsdeletekey

[Run]
Filename: "{{app}}\\{EXE_NAME}.exe"; Description: "{{cm:LaunchProgram,{EXE_NAME}}}"; Flags: nowait postinstall skipifsilent
"""
    with open(ISS_FILE, "w", encoding="utf-8-sig") as f: f.write(iss_content)

    # 3. 编译安装包
    if os.path.exists(ISCC_PATH):
        if subprocess.run(f'"{ISCC_PATH}" "{ISS_FILE}"', shell=True).returncode == 0:
            [shutil.rmtree(f) for f in ['build', 'dist'] if os.path.exists(f)]
            os.remove(f"{EXE_NAME}.spec"); os.remove(ISS_FILE)
            print(f"--- 完成！版本 v{ver} 已就绪 ---")

if __name__ == "__main__": main()