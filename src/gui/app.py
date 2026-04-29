"""Окно customtkinter: выбор файла, обработка, прогресс, лог."""

from __future__ import annotations

import os
import queue
import threading
import traceback
from pathlib import Path

import customtkinter as ctk
import tkinter as tk
import tkinter.messagebox as tkmsg
from customtkinter import CTk, CTkButton, CTkEntry, CTkLabel, CTkProgressBar, CTkTextbox
from loguru import logger
from tkinter import filedialog

from src.core.pipeline import PipelineResult, run_pipeline
from src.utils.logging_setup import setup_logging

APP_TITLE = "XLSX Aggregator"
WINDOW_W, WINDOW_H = 500, 400


def _open_folder_in_explorer(folder: str) -> None:
    """Windows: открыть папку в Проводнике."""
    try:
        os.startfile(folder)  # type: ignore[attr-defined]
    except OSError as e:
        logger.warning("Не удалось открыть папку: %s", e)


def _show_done_with_open_folder(result_path: str) -> None:
    p = Path(result_path)
    folder = str(p.parent)
    msg = f"Готово. Файл обработан: {p.name}\n\nОткрыть папку с файлом?"
    if tkmsg.askyesno("Готово", msg, parent=None):
        _open_folder_in_explorer(folder)


class XlsxAggregatorApp(CTk):
    """Главное окно приложения."""

    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(f"{WINDOW_W}x{WINDOW_H}")
        self.minsize(WINDOW_W, WINDOW_H)
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self._file_path: str = ""
        self._queue: queue.Queue[tuple] = queue.Queue()

        self._path_var = tk.StringVar(value="")
        r0 = ctk.CTkFrame(self)
        r0.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(r0, text="Файл:").pack(side="left", padx=(0, 6))
        self._entry = CTkEntry(r0, textvariable=self._path_var, width=320, state="readonly")
        self._entry.pack(side="left", fill="x", expand=True)
        self._btn_choose = CTkButton(r0, text="Выбрать", width=100, command=self._on_choose)
        self._btn_choose.pack(side="right", padx=(6, 0))

        self._btn_run = CTkButton(self, text="Обработать", width=200, height=36, command=self._on_process)
        self._btn_run.pack(pady=12)
        self._btn_run.configure(state="disabled")

        self._label_pb = ctk.CTkLabel(self, text="0%")
        self._bar = CTkProgressBar(self, width=400)
        self._bar.set(0)
        self._bar.pack(pady=4)
        self._label_pb.pack()
        ctk.CTkLabel(self, text="Лог:").pack(anchor="w", padx=10)
        self._log = CTkTextbox(self, width=460, height=150)
        self._log.pack(padx=10, pady=4, fill="both", expand=True)

        self._status = tk.StringVar(value="Статус: готов")
        ctk.CTkLabel(self, textvariable=self._status).pack(pady=6)
        self.after(100, self._poll_queue)

    def _append_log(self, line: str) -> None:
        self._log.insert("end", line + "\n")
        self._log.see("end")

    def _on_choose(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="Выберите XLSX",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        self._file_path = path
        self._path_var.set(path)
        self._btn_run.configure(state="normal")
        self._append_log(f"Выбран файл: {path}")
        self._status.set("Статус: файл выбран")

    def _on_process(self) -> None:
        if not self._file_path:
            return
        self._btn_run.configure(state="disabled")
        self._btn_choose.configure(state="disabled")
        self._bar.set(0)
        self._status.set("Статус: обработка…")
        t = threading.Thread(target=self._worker, daemon=True)
        t.start()

    def _worker(self) -> None:
        p = self._file_path

        def on_prog(n: int, text: str) -> None:
            self._queue.put(("progress", n, text))
            self._queue.put(("log", f"[{n}%] {text}"))

        try:
            r = run_pipeline(p, progress=on_prog)
            self._queue.put(("ok", r))
        except Exception as e:
            logger.exception("Ошибка пайплайна: {}", e)
            self._queue.put(
                (
                    "err",
                    str(e),
                    traceback.format_exc(),
                )
            )
        finally:
            self._queue.put(("fin", None))

    def _poll_queue(self) -> None:
        try:
            while True:
                item = self._queue.get_nowait()
                if item[0] == "progress":
                    _, n, text = item
                    self._bar.set(max(0.0, min(1.0, n / 100.0)))
                    self._label_pb.configure(text=f"{n}%")
                elif item[0] == "log":
                    self._append_log(item[1])
                elif item[0] == "ok":
                    r: PipelineResult = item[1]
                    self._append_log("Готово.")
                    self._status.set("Статус: готов")
                    self._btn_run.configure(state="normal")
                    self._btn_choose.configure(state="normal")
                    _show_done_with_open_folder(r.result_path)
                elif item[0] == "err":
                    self._append_log(f"Ошибка: {item[1]}\n{item[2]}")
                    self._status.set("Статус: ошибка")
                    self._btn_run.configure(state="normal")
                    self._btn_choose.configure(state="normal")
                    tkmsg.showerror("Ошибка", item[1], parent=self)
                elif item[0] == "fin":
                    break
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)


def run_app() -> None:
    """Инициализирует логи и запускает GUI."""
    setup_logging()
    app = XlsxAggregatorApp()
    app.mainloop()
