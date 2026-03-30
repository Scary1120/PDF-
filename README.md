# 🛠️ PDF 工具箱 (PDF-Toolbox)

一款基于 Python 和 CustomTkinter 开发的轻量级、功能全面的 PDF 处理利器。集成了格式转换、页面管理、尺寸统一及极速压缩功能，并支持 GitHub 远程自动更新。

## ✨ 核心功能

* **全能格式转换**：支持 Word、PPT、Excel 与 PDF 之间的互转（基于 win32com 和 pdf2docx）。
* **页面管理**：可视化预览 PDF，支持页面上移、下移、旋转、删除，并可添加文字水印。
* **尺寸统一**：一键将 PDF 所有页面标准化为 A4 或 A3 尺寸。
* **文件合并**：支持多个 PDF 文件按自定义顺序快速合并。
* **极速压缩**：提供四级压缩方案，显著减小文件体积。
* **右键菜单集成**：通过安装程序可将工具集成至 Windows 右键菜单，实现快速调用。
* **自动更新检测**：程序启动时自动对比 GitHub 最新版本，并引导下载全量安装包。

## 🚀 快速使用

### 普通用户
1.  前往 [Releases](https://github.com/Scary1120/PDF-/releases) 页面。
2.  下载最新的 `PDF工具箱_Setup_vX.X.X.exe`。
3.  运行安装程序，根据向导选择安装路径及是否创建桌面快捷方式。

### 开发者
1.  **克隆仓库**：
    ```bash
    git clone [https://github.com/Scary1120/PDF-.git](https://github.com/Scary1120/PDF-.git)
    ```
2.  **安装依赖**：
    ```bash
    pip install -r requirements.txt
    ```
3.  **运行源码**：
    ```bash
    python PDF工具箱.py
    ```

## 🛠️ 项目构建 (Build)

本项目使用 `AutoBuild.py` 配合 **Inno Setup 6** 实现自动化打包发布：

1.  确认本地已安装 [Inno Setup 6](https://jrsoftware.org/isinfo.php)。
2.  在 `AutoBuild.py` 中配置 `ISCC_PATH` 路径。
3.  运行打包脚本：`python AutoBuild.py`。
4.  脚本会自动递增 `version.txt` 中的版本号并生成安装包。

## 📂 目录结构

```text
PDF-/
├── PDF工具箱.py        # 主程序源代码
├── AutoBuild.py       # 自动化打包与版本管理脚本
├── version.txt        # 当前软件版本号
├── .gitignore         # Git 忽略文件配置
└── README.md          # 项目说明文档
