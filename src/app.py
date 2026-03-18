from __future__ import annotations

import tempfile
import threading
import traceback
from pathlib import Path
from tkinter import Button, Entry, Frame, Label, LabelFrame, StringVar, Tk, filedialog, messagebox

from .ppt_builder import build_ppt
from .preview import generate_previews
from .scanner import ScanError, scan_project


class File2PPTApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("File2PPT")
        self.root.geometry("860x360")
        self.root.resizable(False, False)
        self.root.configure(bg="#f3f3f3")

        self.folder_var = StringVar()
        self.filename_var = StringVar()
        self.status_var = StringVar(value="请选择目录并输入 PPT 文件名。")
        self.is_processing = False
        self._heartbeat_step = 0
        self._heartbeat_base = ""

        self._build_layout()

    def _build_layout(self) -> None:
        title = Label(
            self.root,
            text="功能界面（仅供参考）：",
            font=("Microsoft YaHei", 16, "bold"),
            bg="#f3f3f3",
            anchor="w",
        )
        title.pack(fill="x", padx=28, pady=(18, 10))

        panel = LabelFrame(
            self.root,
            bd=1,
            relief="solid",
            padx=34,
            pady=26,
            bg="#f3f3f3",
        )
        panel.pack(fill="both", expand=True, padx=26, pady=(0, 22))

        Label(
            panel,
            text="PPT 文件名：",
            font=("Microsoft YaHei", 14),
            bg="#f3f3f3",
        ).grid(row=0, column=0, sticky="w", pady=(8, 18))
        self.filename_entry = Entry(
            panel,
            textvariable=self.filename_var,
            width=48,
            font=("Microsoft YaHei", 13),
            relief="solid",
            bd=1,
        )
        self.filename_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=(8, 18), ipady=12)

        Label(
            panel,
            text="文件夹目录：",
            font=("Microsoft YaHei", 14),
            bg="#f3f3f3",
        ).grid(row=1, column=0, sticky="w", pady=(10, 22))
        self.folder_entry = Entry(
            panel,
            textvariable=self.folder_var,
            width=48,
            font=("Microsoft YaHei", 13),
            relief="solid",
            bd=1,
        )
        self.folder_entry.grid(row=1, column=1, sticky="ew", pady=(10, 22), ipady=12)

        browse_button = Button(
            panel,
            text="浏览",
            command=self.select_folder,
            width=8,
            font=("Microsoft YaHei", 12),
        )
        browse_button.grid(row=1, column=2, padx=(22, 0), pady=(10, 22), ipady=8)

        buttons = Frame(panel, bg="#f3f3f3")
        buttons.grid(row=2, column=0, columnspan=3, pady=(6, 16))
        self.generate_button = Button(
            buttons,
            text="确定",
            command=self.start_generation,
            width=10,
            font=("Microsoft YaHei", 12),
        )
        self.generate_button.pack(side="left", padx=(0, 24), ipady=8)

        self.cancel_button = Button(
            buttons,
            text="取消",
            command=self.root.destroy,
            width=10,
            font=("Microsoft YaHei", 12),
        )
        self.cancel_button.pack(side="left", ipady=8)

        status_label = Label(
            panel,
            textvariable=self.status_var,
            anchor="w",
            justify="left",
            wraplength=720,
            fg="#3d4d5c",
            bg="#f3f3f3",
            font=("Microsoft YaHei", 10),
        )
        status_label.grid(row=3, column=0, columnspan=3, sticky="ew")

        panel.columnconfigure(1, weight=1)

    def select_folder(self) -> None:
        selected = filedialog.askdirectory(title="选择总目录")
        if selected:
            self.folder_var.set(selected)

    def start_generation(self) -> None:
        root_dir = self.folder_var.get().strip()
        filename = self.filename_var.get().strip()
        if not root_dir:
            messagebox.showerror("缺少目录", "请先选择文件夹目录。")
            return
        if not filename:
            messagebox.showerror("缺少文件名", "请先输入 PPT 文件名。")
            return

        self.generate_button.configure(state="disabled")
        self.cancel_button.configure(state="disabled")
        self._begin_processing("扫描目录中")
        thread = threading.Thread(
            target=self._run_generation,
            args=(root_dir, filename),
            daemon=True,
        )
        thread.start()

    def _run_generation(self, root_dir: str, filename: str) -> None:
        try:
            project = scan_project(root_dir, self._on_scan_progress)
            file_count = sum(len(topic.items) for section in project.sections for topic in section.topics)
            section_count = len(project.sections)
            self._set_status(f"扫描完成，共发现 {section_count} 组内容、{file_count} 个文件。")
            with tempfile.TemporaryDirectory(prefix="file2ppt_") as preview_dir:
                self._begin_processing("生成预览图中")
                project = generate_previews(project, preview_dir, self._on_preview_progress)
                output_path = Path(root_dir) / f"{filename}.pptx"
                self._begin_processing("写入 PPT 中")
                result = build_ppt(project, output_path, self._on_build_progress)
            self.root.after(0, lambda: self._on_success(result.output_path, result.slide_count, len(result.skipped_files)))
        except ScanError as exc:
            self.root.after(0, lambda: self._on_error(str(exc)))
        except Exception as exc:
            details = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.root.after(0, lambda: self._on_error(f"生成失败: {details}"))

    def _set_status(self, message: str) -> None:
        self.root.after(0, lambda: self.status_var.set(message))

    def _begin_processing(self, base_message: str) -> None:
        self._heartbeat_base = base_message
        self._heartbeat_step = 0
        if not self.is_processing:
            self.is_processing = True
            self.root.after(0, self._heartbeat)
        self._set_status(f"{base_message}...")

    def _heartbeat(self) -> None:
        if not self.is_processing:
            return
        dots = "." * ((self._heartbeat_step % 3) + 1)
        self.status_var.set(f"{self._heartbeat_base}{dots}")
        self._heartbeat_step += 1
        self.root.after(500, self._heartbeat)

    def _on_scan_progress(self, current: int, current_dir: str) -> None:
        self._heartbeat_base = f"扫描目录中，已检查 {current} 个文件夹"
        self._set_status(f"{self._heartbeat_base}：{current_dir}")

    def _on_preview_progress(self, current: int, total: int, filename: str) -> None:
        self._heartbeat_base = f"生成预览图中（{current}/{total}）"
        self._set_status(f"生成预览图中（{current}/{total}）：{filename}")

    def _on_build_progress(self, current: int, total: int, title: str) -> None:
        self._heartbeat_base = f"写入 PPT 中（{current}/{total} 页）"
        self._set_status(f"写入 PPT 中（{current}/{total} 页）：{title}")

    def _on_success(self, output_path: Path, slide_count: int, skipped_count: int) -> None:
        self.is_processing = False
        self.generate_button.configure(state="normal")
        self.cancel_button.configure(state="normal")
        message = f"生成成功，共 {slide_count} 页。输出: {output_path}"
        if skipped_count:
            message += f"；跳过不支持文件 {skipped_count} 个。"
        self.status_var.set(message)
        messagebox.showinfo("生成成功", message)

    def _on_error(self, message: str) -> None:
        self.is_processing = False
        self.generate_button.configure(state="normal")
        self.cancel_button.configure(state="normal")
        self.status_var.set(message)
        messagebox.showerror("生成失败", message)


def main() -> None:
    root = Tk()
    app = File2PPTApp(root)
    app.filename_entry.focus_set()
    root.mainloop()


if __name__ == "__main__":
    main()
