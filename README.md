# File2PPT

把固定目录结构下的图片和 PDF 自动整理成一个 PowerPoint 文件，并提供 Windows 桌面端入口与单文件 `exe` 打包方案。界面按“PPT 文件名 / 文件夹目录 / 浏览 / 确定 / 取消”的最简窗体实现。

## 功能范围

- 输入总目录和 PPT 文件名，自动生成 `.pptx`
- 支持叶子文件类型：`.jpg` `.jpeg` `.png` `.webp` `.pdf`
- 封面页显示总目录名
- 目录页显示一级子目录
- 内容页按“一级目录 -- 二级目录”生成，并自动分页
- 每个文件下方显示去掉后缀后的文件名
- 不支持的文件会跳过，并在结果提示中汇总数量

## 目录规则

总目录下必须是一级子目录，一级子目录下必须是二级子目录，二级子目录下必须放实际文件。

示例：

```text
2026年国际会议/
  ACL/
    计算机视觉/
      样例文件1.webp
      样例文件2.png
  CVPR/
    自然语言处理/
      样例文件1.jpg
      样例文件2.png
```

## 本地运行

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

## Windows 打包

Windows 上安装 Python 后执行：

```bat
build\build_windows.bat
```

输出文件默认位于 `dist\File2PPT.exe`。

## 自动产出 EXE

仓库已经带了 GitHub Actions 工作流 [.github/workflows/build-exe.yml](/Users/550-ml/Project/Python/File2PPT/.github/workflows/build-exe.yml)：

- 推送到 `main` 或 `master` 会自动在 Windows 环境构建
- 也可以在 GitHub 的 `Actions` 页面手动触发
- 构建完成后，在 workflow 的 artifact 中下载 `File2PPT-windows-exe`

如果你当前开发机不是 Windows，推荐直接用这个工作流拿到真正的 `File2PPT.exe`

## 项目结构

- `main.py`：程序入口
- `src/app.py`：Tkinter 界面
- `src/scanner.py`：目录扫描与校验
- `src/preview.py`：图片和 PDF 预览图生成
- `src/ppt_builder.py`：PPT 构建
- `src/models.py`：数据结构

## 已知约束

- 首版只渲染 PDF 首页
- 目录结构不符合要求时会直接报错
- 字体按 `Microsoft YaHei` 设计，Windows 显示效果最佳
