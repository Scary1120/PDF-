import os
import sys
import threading
import fitz
import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
from PIL import Image
import requests
import webbrowser

# ================= 更新配置 =================
CURRENT_VERSION = "1.0.0" 
GITHUB_USER = "Scary1120"      # 你的用户名
GITHUB_REPO = "PDF-"           # 你的仓库名
API_URL = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"
# ===========================================

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def check_for_updates(silent=True):
    def task():
        try:
            # 访问 GitHub API 获取最新版本
            response = requests.get(API_URL, timeout=10)
            if response.status_code == 200:
                data = response.json()
                remote_version = data['tag_name'].lstrip('v')
                if remote_version > CURRENT_VERSION:
                    download_url = data['assets'][0]['browser_download_url']
                    changelog = data.get('body', '无更新说明')
                    if messagebox.askyesno("发现更新", f"新版本: {remote_version}\n内容: {changelog}\n是否下载？"):
                        webbrowser.open(download_url)
                elif not silent:
                    messagebox.showinfo("提示", "当前已是最新版本")
        except:
            if not silent: messagebox.showwarning("警告", "检查更新失败")

    threading.Thread(target=task, daemon=True).start()

# --- 页面类逻辑 (保持你提供的功能不变) ---
class ConvertPage(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        ctk.CTkLabel(self, text="全能格式转换", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        self.mode = ctk.CTkOptionMenu(self, values=["Word 转 PDF", "PDF 转 Word", "PPT 转 PDF", "Excel 转 PDF"])
        self.mode.pack(); self.path = ""
        self.info = ctk.CTkLabel(self, text="未选择", text_color="gray"); self.info.pack()
        ctk.CTkButton(self, text="选择文件", command=self.sel).pack(pady=10)
        ctk.CTkButton(self, text="开始执行", command=self.start).pack(pady=20)
    def set_external_path(self, p): self.path = p; self.info.configure(text=os.path.basename(p))
    def sel(self): self.path = filedialog.askopenfilename(); self.info.configure(text=os.path.basename(self.path) if self.path else "未选择")
    def start(self):
        if self.path: threading.Thread(target=self.work, daemon=True).start()
    def work(self):
        try:
            import win32com.client; from pdf2docx import Converter
            m, p = self.mode.get(), self.path
            out = os.path.splitext(p)[0] + (".docx" if "PDF 转 Word" == m else ".pdf")
            if m == "PDF 转 Word": cv = Converter(p); cv.convert(out); cv.close()
            elif m == "Word 转 PDF": w = win32com.client.Dispatch("Word.Application"); d = w.Documents.Open(p); d.SaveAs(out, 17); d.Close(); w.Quit()
            elif m == "PPT 转 PDF": ppt = win32com.client.Dispatch("Powerpoint.Application"); d = ppt.Presentations.Open(p, WithWindow=False); d.SaveAs(out, 2); d.Close(); ppt.Quit()
            elif m == "Excel 转 PDF": e = win32com.client.Dispatch("Excel.Application"); d = e.Workbooks.Open(p); d.ExportAsFixedFormat(0, out); d.Close(); e.Quit()
            messagebox.showinfo("OK", "转换成功")
        except Exception as e: messagebox.showerror("失败", str(e))

class PageManagePage(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1); self.grid_columnconfigure(1, weight=1); self.grid_rowconfigure(0, weight=1)
        self.L = ctk.CTkFrame(self); self.L.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.path = ""; self.doc = None; self.rots = {}
        ctk.CTkButton(self.L, text="加载 PDF", command=self.load).pack(pady=10)
        self.lb = tk.Listbox(self.L, height=18); self.lb.pack(fill="both", expand=True, padx=20); self.lb.bind("<<ListboxSelect>>", self.pre)
        btns = ctk.CTkFrame(self.L, fg_color="transparent"); btns.pack(pady=5)
        ctk.CTkButton(btns, text="上移", width=50, command=self.up).pack(side="left", padx=2)
        ctk.CTkButton(btns, text="下移", width=50, command=self.dn).pack(side="left", padx=2)
        ctk.CTkButton(btns, text="旋转", width=50, fg_color="#F39C12", command=self.rot).pack(side="left", padx=2)
        ctk.CTkButton(btns, text="删除", width=50, fg_color="red", command=self.rm).pack(side="left", padx=2)
        self.wm = ctk.CTkEntry(self.L, placeholder_text="水印文字"); self.wm.pack(pady=5, padx=20, fill="x")
        ctk.CTkButton(self.L, text="保存修改", fg_color="green", command=self.sv).pack(pady=10)
        self.R = ctk.CTkFrame(self); self.R.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.pre_l = ctk.CTkLabel(self.R, text="预览区"); self.pre_l.pack(expand=True)
    def load(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if p:
            self.path = p; self.doc = fitz.open(p); self.lb.delete(0, tk.END); self.rots = {}
            for i in range(len(self.doc)): self.lb.insert(tk.END, f"第 {i+1} 页"); self.rots[i] = 0
    def pre(self, e):
        if not self.doc or not self.lb.curselection(): return
        idx = int(self.lb.get(self.lb.curselection()[0]).split(" ")[1]) - 1
        pix = self.doc[idx].get_pixmap(matrix=fitz.Matrix(0.4, 0.4))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        if self.rots.get(idx, 0) != 0: img = img.rotate(-self.rots[idx], expand=True)
        img.thumbnail((350, 450)); ctk_img = ctk.CTkImage(img, size=(img.width, img.height)); self.pre_l.configure(image=ctk_img, text="")
    def up(self):
        i = self.lb.curselection()
        if i and i[0] > 0: v = self.lb.get(i); self.lb.delete(i); self.lb.insert(i[0]-1, v); self.lb.select_set(i[0]-1)
    def dn(self):
        i = self.lb.curselection()
        if i and i[0] < self.lb.size()-1: v = self.lb.get(i); self.lb.delete(i); self.lb.insert(i[0]+1, v); self.lb.select_set(i[0]+1)
    def rot(self):
        if not self.lb.curselection(): return
        idx = int(self.lb.get(self.lb.curselection()[0]).split(" ")[1]) - 1
        self.rots[idx] = (self.rots[idx] + 90) % 360; self.pre(None)
    def rm(self):
        if self.lb.curselection(): self.lb.delete(self.lb.curselection())
    def sv(self):
        out = filedialog.asksaveasfilename(defaultextension=".pdf")
        if out:
            nd = fitz.open(); t = self.wm.get()
            for item in self.lb.get(0, tk.END):
                ox = int(item.split(" ")[1]) - 1
                nd.insert_pdf(self.doc, from_page=ox, to_page=ox)
                pg = nd[-1]; pg.set_rotation(self.rots[ox])
                if t: pg.insert_text((50, 50), t, fontsize=40, color=(0.8, 0.8, 0.8), rotate=45)
            nd.save(out); nd.close(); messagebox.showinfo("OK", "导出成功")

class ResizePage(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.p = ""
        ctk.CTkLabel(self, text="尺寸统一", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        ctk.CTkButton(self, text="选择 PDF", command=self.sel).pack(pady=10)
        self.sz = ctk.CTkOptionMenu(self, values=["A4", "A3"]); self.sz.pack(pady=10); self.sz.set("A4")
        ctk.CTkButton(self, text="执行", fg_color="green", command=self.go).pack(pady=20)
    def sel(self): self.p = filedialog.askopenfilename()
    def go(self):
        tw = 595 if self.sz.get()=="A4" else 842; d = fitz.open(self.p); nd = fitz.open()
        for pg in d: th = pg.rect.height * (tw / pg.rect.width); nd.new_page(width=tw, height=th).show_pdf_page(fitz.Rect(0,0,tw,th), d, pg.number)
        nd.save(self.p.replace(".pdf", "_sz.pdf")); nd.close(); d.close(); messagebox.showinfo("OK", "完成")

class MergePage(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.fs = []; self.lb = tk.Listbox(self, height=12); self.lb.pack(fill="both", padx=40, pady=10)
        ctk.CTkButton(self, text="添加文件", command=self.add).pack()
        ctk.CTkButton(self, text="合并保存", fg_color="green", command=self.go).pack(pady=20)
    def add(self): ns = filedialog.askopenfilenames(); [ (self.fs.append(n), self.lb.insert(tk.END, os.path.basename(n))) for n in ns ]
    def go(self):
        from PyPDF2 import PdfMerger
        m = PdfMerger(); [ m.append(f) for f in self.fs ]; out = filedialog.asksaveasfilename(defaultextension=".pdf")
        if out: m.write(out); m.close(); messagebox.showinfo("OK", "成功")

class CompressPage(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.p = ""
        ctk.CTkLabel(self, text="PDF 极速压缩", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        ctk.CTkButton(self, text="选择 PDF", command=self.sel).pack(pady=10)
        self.slider = ctk.CTkSlider(self, from_=1, to=4, number_of_steps=3); self.slider.set(2); self.slider.pack(pady=20)
        ctk.CTkButton(self, text="开始压缩", fg_color="#e67e22", command=self.go).pack()
    def sel(self): self.p = filedialog.askopenfilename()
    def go(self):
        if self.p: doc = fitz.open(self.p); doc.save(self.p.replace(".pdf","_min.pdf"), garbage=int(self.slider.get()), deflate=True, clean=True); doc.close(); messagebox.showinfo("OK", "完成")

# --- 主窗口逻辑 ---
class PDFToolBox(ctk.CTk):
    def __init__(self, external_file=None):
        super().__init__()
        self.title(f"PDF工具箱 v{CURRENT_VERSION}")
        self.geometry("1100x750")
        self.grid_columnconfigure(1, weight=1); self.grid_rowconfigure(0, weight=1)
        self.nav_frame = ctk.CTkFrame(self, corner_radius=0); self.nav_frame.grid(row=0, column=0, sticky="nsew")
        ctk.CTkLabel(self.nav_frame, text="PDF工具箱", font=ctk.CTkFont(size=22, weight="bold")).pack(pady=35)
        
        self.menu_items = [("格式转换", ConvertPage), ("页面管理", PageManagePage), ("尺寸统一", ResizePage), ("文件合并", MergePage), ("PDF压缩", CompressPage)]
        for text, page_class in self.menu_items:
            ctk.CTkButton(self.nav_frame, text=text, corner_radius=0, height=45, fg_color="transparent", anchor="w", command=lambda p=page_class: self.switch_page(p)).pack(fill="x")
            
        ctk.CTkButton(self.nav_frame, text="检查更新", fg_color="gray25", height=30, command=lambda: check_for_updates(silent=False)).pack(side="bottom", pady=20, padx=10)
        ctk.CTkLabel(self.nav_frame, text=f"版本: {CURRENT_VERSION}", font=ctk.CTkFont(size=10), text_color="gray").pack(side="bottom")
            
        self.container = ctk.CTkFrame(self, fg_color="transparent"); self.container.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.current_page = None; self.switch_page(ConvertPage)
        self.after(1000, lambda: check_for_updates(silent=True)) # 启动检测
        if external_file: self.after(200, lambda: self.current_page.set_external_path(external_file))
        
    def switch_page(self, page_class):
        if self.current_page: self.current_page.destroy()
        self.current_page = page_class(self.container); self.current_page.pack(expand=True, fill="both")

if __name__ == "__main__":
    app = PDFToolBox(external_file=sys.argv[1] if len(sys.argv) > 1 else None)
    app.mainloop()