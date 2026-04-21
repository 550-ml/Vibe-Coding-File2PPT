from __future__ import annotations

import tempfile
import threading
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import Button, Entry, Frame, Label, LabelFrame, StringVar, Tk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from .control import check_software_control
from .ppt_builder import build_ppt
from .preview import generate_previews
from .scanner import ScanError, scan_project


class File2PPTApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("File2PPT")
        self.root.geometry("980x620")
        self.root.minsize(900, 520)
        self.root.resizable(True, True)
        self.root.configure(bg="#f3f3f3")

        self.folder_var = StringVar()
        self.filename_var = StringVar()
        self.status_var = StringVar(value="请选择目录并输入 PPT 文件名。")
        self.output_hint_var = StringVar(value="输出位置：将保存到所选目录的上一级目录。")
        self.is_processing = False
        self._heartbeat_step = 0
        self._heartbeat_base = ""
        self._last_logged_message = ""

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
            pady=24,
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

        output_hint_label = Label(
            panel,
            textvariable=self.output_hint_var,
            anchor="w",
            justify="left",
            fg="#51606f",
            bg="#f3f3f3",
            font=("Microsoft YaHei", 10),
        )
        output_hint_label.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(4, 8))

        status_label = Label(
            panel,
            textvariable=self.status_var,
            anchor="w",
            justify="left",
            wraplength=820,
            fg="#304355",
            bg="#f3f3f3",
            font=("Microsoft YaHei", 10, "bold"),
        )
        status_label.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 8))

        log_label = Label(
            panel,
            text="运行日志：",
            font=("Microsoft YaHei", 11, "bold"),
            bg="#f3f3f3",
            anchor="w",
        )
        log_label.grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 6))

        self.log_text = ScrolledText(
            panel,
            height=12,
            wrap="word",
            font=("Microsoft YaHei", 10),
            relief="solid",
            bd=1,
        )
        self.log_text.grid(row=6, column=0, columnspan=3, sticky="nsew")
        self.log_text.configure(state="disabled")

        panel.columnconfigure(1, weight=1)
        panel.rowconfigure(6, weight=1)

    def select_folder(self) -> None:
        selected = filedialog.askdirectory(title="选择总目录")
        if selected:
            self.folder_var.set(selected)
            if not self.filename_var.get().strip():
                self.filename_var.set(Path(selected).name)
            output_path = self._resolve_output_path(selected, self.filename_var.get().strip() or Path(selected).name)
            self.output_hint_var.set(f"输出位置：{output_path}")
            self._append_log(f"已选择目录：{selected}")

    def start_generation(self) -> None:
        root_dir = self.folder_var.get().strip()
        filename = self.filename_var.get().strip() or Path(root_dir).name
        if not root_dir:
            messagebox.showerror("缺少目录", "请先选择文件夹目录。")
            return

        self.filename_var.set(filename)
        output_path = self._resolve_output_path(root_dir, filename)

        self.generate_button.configure(state="disabled")
        self.cancel_button.configure(state="disabled")
        self._append_log(f"准备生成 PPT：{output_path}")
        self._begin_processing("扫描目录中")
        thread = threading.Thread(
            target=self._run_generation,
            args=(root_dir, filename, output_path),
            daemon=True,
        )
        thread.start()

    def _run_generation(self, root_dir: str, filename: str, output_path: Path) -> None:
        try:
            project = scan_project(root_dir, self._on_scan_progress)
            file_count = sum(len(topic.items) for section in project.sections for topic in section.topics)
            section_count = len(project.sections)
            self._report_status(f"扫描完成，共发现 {section_count} 组内容、{file_count} 个文件。")
            with tempfile.TemporaryDirectory(prefix="file2ppt_") as preview_dir:
                self._begin_processing("生成预览图中")
                project = generate_previews(project, preview_dir, self._on_preview_progress)
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

    def _report_status(self, message: str) -> None:
        self.root.after(0, lambda: self._apply_status_and_log(message))

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
        self._report_status(f"{self._heartbeat_base}：{current_dir}")

    def _on_preview_progress(self, current: int, total: int, filename: str) -> None:
        self._heartbeat_base = f"生成预览图中（{current}/{total}）"
        self._report_status(f"生成预览图中（{current}/{total}）：{filename}")

    def _on_build_progress(self, current: int, total: int, title: str) -> None:
        self._heartbeat_base = f"写入 PPT 中（{current}/{total} 页）"
        self._report_status(f"写入 PPT 中（{current}/{total} 页）：{title}")

    def _on_success(self, output_path: Path, slide_count: int, skipped_count: int) -> None:
        self.is_processing = False
        self.generate_button.configure(state="normal")
        self.cancel_button.configure(state="normal")
        message = f"生成成功，共 {slide_count} 页。输出: {output_path}"
        if skipped_count:
            message += f"；跳过不支持文件 {skipped_count} 个。"
        self._apply_status_and_log(message)
        messagebox.showinfo("生成成功", message)

    def _on_error(self, message: str) -> None:
        self.is_processing = False
        self.generate_button.configure(state="normal")
        self.cancel_button.configure(state="normal")
        self._apply_status_and_log(message)
        messagebox.showerror("生成失败", message)

    def _apply_status_and_log(self, message: str) -> None:
        self.status_var.set(message)
        self._append_log(message)

    def _append_log(self, message: str) -> None:
        if message == self._last_logged_message:
            return
        self._last_logged_message = message
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _resolve_output_path(self, root_dir: str, filename: str) -> Path:
        root_path = Path(root_dir).expanduser().resolve()
        safe_name = filename.strip() or root_path.name
        return root_path.parent / f"{safe_name}.pptx"


def main() -> None:
    root = Tk()
    root.withdraw()
    control_result = check_software_control()
    if not control_result.ok:
        messagebox.showerror("授权校验失败", control_result.message)
        root.destroy()
        return
    root.deiconify()
    app = File2PPTApp(root)
    app.filename_entry.focus_set()
    root.mainloop()


if __name__ == "__main__":
    main()
