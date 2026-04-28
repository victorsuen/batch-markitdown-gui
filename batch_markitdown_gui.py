from __future__ import annotations

import json
import re
import threading
import traceback
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, TimeoutError
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import numpy as np
import pypdfium2 as pdfium
from markitdown import MarkItDown
from rapidocr_onnxruntime import RapidOCR
from tkinterdnd2 import DND_FILES, TkinterDnD

from batch_markitdown import SUPPORTED_EXTENSIONS, discover_files, make_output_path

APP_VERSION = "v1.10"
CONFIG_PATH = Path.home() / ".batch_markitdown_gui.json"


class BatchMarkItDownApp:
    def __init__(self, root: TkinterDnD.Tk) -> None:
        self.root = root
        self.root.title("Batch MarkItDown Converter")
        self.root.geometry("860x560")
        self.root.minsize(760, 480)

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.remember_output_dir = tk.BooleanVar(value=True)
        self.recursive = tk.BooleanVar(value=True)
        self.use_plugins = tk.BooleanVar(value=False)
        self.fast_mode = tk.BooleanVar(value=True)
        self.financial_table_mode = tk.BooleanVar(value=True)
        self.timeout_seconds = tk.IntVar(value=120)
        self.status_text = tk.StringVar(value="Ready")
        self.drop_info_text = tk.StringVar(value="尚未拖入檔案")
        self.dropped_files: list[Path] = []
        self.ocr_engine: RapidOCR | None = None

        self._load_config()
        self._apply_styles()
        self._build_ui()

    def _apply_styles(self) -> None:
        self.root.configure(bg="#eef3fb")
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Banner.TFrame", background="#2f66d0")
        style.configure("Banner.TLabel", background="#2f66d0", foreground="#ffffff")
        style.configure("Title.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("Hint.TLabel", foreground="#4d5b78", background="#ffffff")
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

    def _load_config(self) -> None:
        if not CONFIG_PATH.exists():
            return
        try:
            raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            remembered = raw.get("output_dir")
            remember_enabled = bool(raw.get("remember_output_dir", True))
            self.remember_output_dir.set(remember_enabled)
            if remember_enabled and isinstance(remembered, str) and remembered.strip():
                self.output_dir.set(remembered)
            timeout_val = int(raw.get("timeout_seconds", 120))
            self.timeout_seconds.set(max(30, min(timeout_val, 1800)))
            self.fast_mode.set(bool(raw.get("fast_mode", True)))
            self.financial_table_mode.set(bool(raw.get("financial_table_mode", True)))
        except Exception:
            # Ignore invalid config to keep app usable.
            return

    def _save_config(self) -> None:
        if not self.remember_output_dir.get():
            payload = {"remember_output_dir": False}
        else:
            payload = {
                "remember_output_dir": True,
                "output_dir": self.output_dir.get().strip(),
                "timeout_seconds": int(self.timeout_seconds.get()),
                "fast_mode": bool(self.fast_mode.get()),
                "financial_table_mode": bool(self.financial_table_mode.get()),
            }
        try:
            CONFIG_PATH.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            return

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=14)
        main.pack(fill=tk.BOTH, expand=True)

        banner = ttk.Frame(main, style="Banner.TFrame", padding=(12, 8))
        banner.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Label(
            banner, text=f"Batch MarkItDown Converter {APP_VERSION}", style="Banner.TLabel"
        ).pack(side=tk.LEFT)
        ttk.Label(
            banner, text="批量轉換文件為 Markdown", style="Banner.TLabel"
        ).pack(side=tk.RIGHT)

        card = ttk.Frame(main, style="Card.TFrame", padding=12)
        card.grid(row=1, column=0, columnspan=2, sticky="nsew")

        ttk.Label(card, text="來源資料夾", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(card, textvariable=self.input_dir).grid(
            row=1, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Button(card, text="選擇...", command=self.pick_input_dir).grid(
            row=1, column=1, sticky="ew"
        )

        ttk.Label(card, text="輸出資料夾", style="Title.TLabel").grid(
            row=2, column=0, sticky="w", pady=(12, 0)
        )
        ttk.Entry(card, textvariable=self.output_dir).grid(
            row=3, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Button(card, text="選擇...", command=self.pick_output_dir).grid(
            row=3, column=1, sticky="ew"
        )
        ttk.Checkbutton(
            card,
            text="記錄輸出資料夾（下次開啟自動沿用）",
            variable=self.remember_output_dir,
            command=self._save_config,
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        drop_frame = ttk.Frame(card)
        drop_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        drop_label = tk.Label(
            drop_frame,
            text="拖放區\n直接拖入資料夾或檔案\n（支援副檔名）",
            relief="ridge",
            bd=2,
            padx=14,
            pady=20,
            anchor="center",
            justify="center",
            bg="#f2f7ff",
            fg="#1f3d76",
            font=("Segoe UI", 10, "bold"),
            height=6,
        )
        drop_label.pack(fill=tk.X, expand=True)
        drop_label.drop_target_register(DND_FILES)
        drop_label.dnd_bind("<<Drop>>", self.on_drop)
        ttk.Label(
            card,
            textvariable=self.drop_info_text,
            style="Hint.TLabel",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))

        options = ttk.Frame(card, style="Card.TFrame")
        options.grid(row=6, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Checkbutton(options, text="遞迴掃描子資料夾", variable=self.recursive).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Checkbutton(options, text="啟用 plugins", variable=self.use_plugins).pack(
            side=tk.LEFT
        )
        ttk.Checkbutton(options, text="快速模式（convert_local）", variable=self.fast_mode).pack(
            side=tk.LEFT, padx=(12, 0)
        )
        ttk.Checkbutton(
            options, text="財務表格模式", variable=self.financial_table_mode
        ).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Label(options, text="單檔逾時(秒):").pack(side=tk.LEFT, padx=(14, 4))
        timeout_spin = ttk.Spinbox(
            options,
            from_=30,
            to=1800,
            textvariable=self.timeout_seconds,
            width=6,
            increment=10,
        )
        timeout_spin.pack(side=tk.LEFT)

        action_row = ttk.Frame(card, style="Card.TFrame")
        action_row.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.start_button = ttk.Button(
            action_row, text="開始轉換", style="Accent.TButton", command=self.start_convert
        )
        self.start_button.pack(side=tk.LEFT)
        ttk.Button(action_row, text="清空紀錄", command=self.clear_log).pack(
            side=tk.LEFT, padx=(8, 0)
        )

        self.progress = ttk.Progressbar(card, mode="determinate")
        self.progress.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        ttk.Label(card, textvariable=self.status_text, style="Hint.TLabel").grid(
            row=9, column=0, columnspan=2, sticky="w", pady=(8, 4)
        )

        self.log = tk.Text(
            card,
            wrap="word",
            bg="#fbfdff",
            fg="#1f1f1f",
            insertbackground="#1f1f1f",
            relief="solid",
            bd=1,
            font=("Consolas", 10),
        )
        self.log.grid(row=10, column=0, columnspan=2, sticky="nsew")

        scroll = ttk.Scrollbar(card, orient="vertical", command=self.log.yview)
        scroll.grid(row=10, column=2, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)

        ttk.Label(
            card,
            text="支援副檔名: " + ", ".join(sorted(SUPPORTED_EXTENSIONS)),
            style="Hint.TLabel",
        ).grid(row=11, column=0, sticky="w", pady=(8, 0))
        ttk.Label(card, text=f"版本: {APP_VERSION}", style="Hint.TLabel").grid(
            row=11, column=1, sticky="e", pady=(8, 0)
        )

        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(10, weight=1)

    def on_drop(self, event: tk.Event) -> None:
        raw_paths = self.root.tk.splitlist(event.data)
        dropped = [Path(item) for item in raw_paths if item]
        if not dropped:
            return

        folders = [p for p in dropped if p.exists() and p.is_dir()]
        files = [p for p in dropped if p.exists() and p.is_file()]

        if folders:
            folder = folders[0].resolve()
            self.input_dir.set(str(folder))
            self.dropped_files = []
            self.drop_info_text.set(f"已選資料夾：{folder.name}")
            self.append_log(f"已設定來源資料夾（拖放）: {folder}")
            self.status_text.set("已設定來源資料夾")
            return

        supported_files = [
            p.resolve() for p in files if p.suffix.lower() in SUPPORTED_EXTENSIONS
        ]
        if not supported_files:
            messagebox.showwarning("提示", "拖入的檔案沒有支援的副檔名。")
            return

        self.dropped_files = supported_files
        self.input_dir.set("")
        names = [p.name for p in supported_files]
        preview = "、".join(names[:4])
        if len(names) > 4:
            preview += f" ...（共 {len(names)} 個）"
        self.drop_info_text.set(f"已拖入檔案：{preview}")
        self.append_log(f"已加入拖放檔案 {len(supported_files)} 個。")
        for p in supported_files[:12]:
            self.append_log(f"  - {p.name}")
        if len(supported_files) > 12:
            self.append_log(f"  ...其餘 {len(supported_files) - 12} 個已省略")
        self.status_text.set(f"已加入拖放檔案 {len(supported_files)} 個")

    def pick_input_dir(self) -> None:
        selected = filedialog.askdirectory(title="選擇來源資料夾")
        if selected:
            self.input_dir.set(selected)

    def pick_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="選擇輸出資料夾")
        if selected:
            self.output_dir.set(selected)
            self._save_config()

    def append_log(self, text: str) -> None:
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)

    def clear_log(self) -> None:
        self.log.delete("1.0", tk.END)

    def start_convert(self) -> None:
        src_text = self.input_dir.get().strip()
        dst = Path(self.output_dir.get().strip()).expanduser()

        if not src_text and not self.dropped_files:
            messagebox.showerror("錯誤", "請先選擇來源資料夾，或拖放檔案/資料夾。")
            return

        src = Path(src_text).expanduser() if src_text else None
        if src is not None and (not src.exists() or not src.is_dir()):
            messagebox.showerror("錯誤", "請先選擇有效的來源資料夾。")
            return

        dst.mkdir(parents=True, exist_ok=True)
        self._save_config()

        self.start_button.configure(state="disabled")
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)
        self.progress["value"] = 0
        self.status_text.set("掃描檔案中...")
        if src is not None:
            self.append_log(f"來源資料夾: {src}")
        elif self.dropped_files:
            self.append_log(f"來源檔案（拖放）: {len(self.dropped_files)} 個")
        self.append_log(f"輸出: {dst}")

        worker = threading.Thread(
            target=self._convert_worker,
            args=(src.resolve() if src is not None else None, dst.resolve()),
            daemon=True,
        )
        worker.start()

    def _unique_output_for_file(self, file_path: Path, output_dir: Path) -> Path:
        candidate = (output_dir / file_path.name).with_suffix(".md")
        if not candidate.exists():
            return candidate

        idx = 1
        while True:
            renamed = candidate.with_name(f"{candidate.stem}_{idx}.md")
            if not renamed.exists():
                return renamed
            idx += 1

    def _convert_worker(self, src: Path | None, dst: Path) -> None:
        try:
            files = (
                discover_files(src, recursive=self.recursive.get())
                if src is not None
                else self.dropped_files
            )
            if not files:
                self.root.after(0, self._finish_empty)
                return

            total = len(files)
            ok_count = 0
            fail_count = 0

            self.root.after(0, lambda: self.status_text.set(f"掃描完成，共 {total} 個檔案"))
            self.root.after(0, lambda: self.append_log(f"掃描完成，共 {total} 個檔案"))
            self.root.after(0, lambda: self.status_text.set("初始化轉換器中..."))
            converter = MarkItDown(enable_plugins=self.use_plugins.get())
            timeout_seconds = max(30, int(self.timeout_seconds.get()))
            use_fast_mode = bool(self.fast_mode.get())

            self.root.after(0, self.progress.stop)
            self.root.after(
                0, lambda: self.progress.configure(mode="determinate", maximum=total, value=0)
            )
            self.root.after(
                0,
                lambda: self.status_text.set(
                    f"開始轉換，共 {total} 個檔案...（逾時 {timeout_seconds}s）"
                ),
            )

            for idx, file_path in enumerate(files, start=1):
                self.root.after(
                    0,
                    lambda n=idx, t=total, p=file_path: self.status_text.set(
                        f"轉換中... {n}/{t} | {p.name}"
                    ),
                )
                if src is None:
                    output_md = self._unique_output_for_file(file_path, dst)
                else:
                    output_md = make_output_path(file_path, src, dst)
                output_md.parent.mkdir(parents=True, exist_ok=True)
                try:
                    with ThreadPoolExecutor(max_workers=1) as pool:
                        fn = converter.convert_local if use_fast_mode else converter.convert
                        future = pool.submit(fn, str(file_path))
                        result = future.result(timeout=timeout_seconds)
                    content = result.text_content or ""
                    if file_path.suffix.lower() == ".pdf" and len(content.strip()) < 20:
                        self.root.after(
                            0,
                            lambda p=file_path: self.append_log(
                                f"[OCR] {p}: 文字層近乎空白，改用 OCR..."
                            ),
                        )
                        ocr_content = self._ocr_pdf_to_markdown(file_path)
                        if ocr_content.strip():
                            content = ocr_content
                            self.root.after(
                                0,
                                lambda p=file_path: self.append_log(
                                    f"[OCR-OK] {p}: OCR 完成"
                                ),
                            )
                        else:
                            self.root.after(
                                0,
                                lambda p=file_path: self.append_log(
                                    f"[OCR-EMPTY] {p}: OCR 仍無可用文字"
                                ),
                            )

                    output_md.write_text(content, encoding="utf-8")
                    ok_count += 1
                    self.root.after(
                        0, lambda p=file_path, o=output_md: self.append_log(f"[OK] {p} -> {o}")
                    )
                except TimeoutError:
                    fail_count += 1
                    self.root.after(
                        0,
                        lambda p=file_path, t=timeout_seconds: self.append_log(
                            f"[TIMEOUT] {p}: 超過 {t} 秒，已跳過"
                        ),
                    )
                except Exception as exc:  # noqa: BLE001
                    fail_count += 1
                    self.root.after(0, lambda p=file_path, e=exc: self.append_log(f"[FAIL] {p}: {e}"))

                self.root.after(0, lambda n=idx: self.progress.configure(value=n))

            self.root.after(0, lambda: self._finish_done(ok_count, fail_count, dst))
        except Exception as exc:  # noqa: BLE001
            detail = traceback.format_exc()
            self.root.after(0, lambda: self._finish_error(exc, detail))

    def _get_ocr_engine(self) -> RapidOCR:
        if self.ocr_engine is None:
            self.ocr_engine = RapidOCR()
        return self.ocr_engine

    def _ocr_pdf_to_markdown(self, pdf_path: Path) -> str:
        engine = self._get_ocr_engine()
        lines: list[str] = [f"# OCR Result: {pdf_path.name}", ""]
        doc = pdfium.PdfDocument(str(pdf_path))

        for idx in range(len(doc)):
            page = doc[idx]
            bitmap = page.render(scale=2.0)
            pil_image = bitmap.to_pil()
            arr = np.array(pil_image)
            result, _ = engine(arr)

            page_lines: list[str] = []
            if result:
                for item in result:
                    if len(item) < 2:
                        continue
                    txt = item[1]
                    if isinstance(txt, tuple):
                        txt = txt[0]
                    text_value = str(txt).strip()
                    if text_value:
                        page_lines.append(text_value)

            page_lines = self._format_ocr_lines(page_lines)
            if self.financial_table_mode.get():
                page_lines = self._format_financial_rows(page_lines)

            if page_lines:
                lines.append(f"## Page {idx + 1}")
                lines.extend(page_lines)
                lines.append("")

        return "\n".join(lines).strip() + "\n"

    def _is_sentence_end(self, text: str) -> bool:
        return bool(re.search(r"[。！？；.!?;:：)\]）】」』]$", text))

    def _normalize_zh_text(self, text: str) -> str:
        s = text
        # Full-width punctuation normalization for Traditional Chinese readability.
        s = s.replace(",", "，").replace(";", "；").replace(":", "：")
        s = s.replace("(", "（").replace(")", "）")
        # Keep decimal numbers safe, then convert remaining periods to full stop.
        s = re.sub(r"(\d)\.(\d)", r"\1__DECIMAL_DOT__\2", s)
        s = s.replace(".", "。")
        s = s.replace("__DECIMAL_DOT__", ".")
        # Remove extra spaces between CJK and ASCII/alnum.
        s = re.sub(r"([\u4e00-\u9fff])\s+([A-Za-z0-9])", r"\1\2", s)
        s = re.sub(r"([A-Za-z0-9])\s+([\u4e00-\u9fff])", r"\1\2", s)
        # Avoid punctuation leading/trailing space artifacts.
        s = re.sub(r"\s+([，。；：！？）】」』])", r"\1", s)
        s = re.sub(r"([（【「『])\s+", r"\1", s)
        return s.strip()

    def _format_ocr_lines(self, raw_lines: list[str]) -> list[str]:
        cleaned: list[str] = []
        for line in raw_lines:
            # Normalize spaces and remove obvious OCR noise-only lines.
            s = re.sub(r"\s+", " ", line).strip()
            if not s:
                continue
            if re.fullmatch(r"[-_=~·•\s]+", s):
                continue
            cleaned.append(self._normalize_zh_text(s))

        if not cleaned:
            return []

        merged: list[str] = []
        for s in cleaned:
            # Keep table-like lines and list lines as individual lines.
            if "|" in s or re.match(r"^(\d+[.)]|[-*•])\s+", s):
                merged.append(s)
                continue

            if not merged:
                merged.append(s)
                continue

            prev = merged[-1]
            # If previous line likely continues, merge for paragraph readability.
            # Prefer merging CJK lines unless previous line clearly ended.
            zh_heavy_prev = len(re.findall(r"[\u4e00-\u9fff]", prev)) >= max(2, len(prev) // 4)
            zh_heavy_cur = len(re.findall(r"[\u4e00-\u9fff]", s)) >= max(2, len(s) // 4)
            should_merge = not self._is_sentence_end(prev)
            if zh_heavy_prev and zh_heavy_cur and len(prev) < 180:
                should_merge = True
            if should_merge:
                joiner = "" if (zh_heavy_prev and zh_heavy_cur) else " "
                merged[-1] = f"{prev}{joiner}{s}"
            else:
                merged.append(s)

        # Break long merged blocks into readable paragraph chunks.
        output: list[str] = []
        for s in merged:
            if len(s) <= 220:
                output.append(s)
                continue

            buf = s
            while len(buf) > 220:
                split_idx = max(buf.rfind("。", 0, 220), buf.rfind(".", 0, 220), buf.rfind(" ", 0, 220))
                if split_idx < 80:
                    split_idx = 220
                output.append(buf[: split_idx + 1].strip())
                buf = buf[split_idx + 1 :].strip()
            if buf:
                output.append(buf)

        return output

    def _looks_like_financial_row(self, text: str) -> bool:
        number_tokens = re.findall(r"-?\d[\d,]*(?:\.\d+)?%?", text)
        has_gap_columns = bool(re.search(r"\s{2,}", text))
        has_money_mark = any(ch in text for ch in ("$", "HK$", "RMB", "USD", "港元", "元", "%"))
        return len(number_tokens) >= 2 and (has_gap_columns or has_money_mark)

    def _format_financial_rows(self, lines: list[str]) -> list[str]:
        if not lines:
            return lines

        out: list[str] = []
        for s in lines:
            if self._looks_like_financial_row(s):
                # Keep multi-space separation so columns remain visually distinct in markdown.
                s_norm = re.sub(r"\s{3,}", "  |  ", s.strip())
                if "|" not in s_norm:
                    s_norm = re.sub(r"\s{2,}", "  |  ", s_norm)
                out.append(s_norm)
            else:
                out.append(s)
        return out

    def _finish_empty(self) -> None:
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.start_button.configure(state="normal")
        self.status_text.set("找不到可轉換檔案")
        self.append_log("找不到支援的檔案。")
        messagebox.showinfo("完成", "找不到支援的檔案。")

    def _finish_done(self, ok_count: int, fail_count: int, dst: Path) -> None:
        self.progress.stop()
        self.start_button.configure(state="normal")
        self.status_text.set(f"完成，成功 {ok_count}，失敗 {fail_count}")
        self.append_log("")
        self.append_log(f"完成。成功: {ok_count}, 失敗: {fail_count}")
        self.append_log(f"輸出位置: {dst}")
        messagebox.showinfo("完成", f"成功: {ok_count}\n失敗: {fail_count}\n\n輸出: {dst}")

    def _finish_error(self, exc: Exception, detail: str) -> None:
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.start_button.configure(state="normal")
        self.status_text.set("執行失敗，請查看紀錄")
        self.append_log(f"[ERROR] 程式錯誤: {exc}")
        self.append_log(detail)
        messagebox.showerror("錯誤", f"程式執行失敗：{exc}\n請查看紀錄區。")


def main() -> None:
    root = TkinterDnD.Tk()
    app = BatchMarkItDownApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
