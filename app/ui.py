"""customtkinter ile masaüstü arayüzü: dosya seç/sürükle, hedef format seç,
dönüştür. Tüm iş controller.py üzerinden arka plan thread'inde yapılır; bu
modül sadece queue.Queue'yu root.after() ile periyodik okuyup UI'yi günceller.

Görsel dil: sıcak/krem editoryal tema, düz (corner_radius=0) yüzeyler,
tek vurgu rengi (mercan/turuncu). Bilinçli olarak tipik "AI aracı" koyu
mavi/gradient temasından kaçınılmıştır.
"""
from __future__ import annotations

import logging
import os
import queue
import traceback
from pathlib import Path

import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD

from . import controller, security
from .controller import ConversionJob
from .resources import resource_path

# ---------------------------------------------------------------- palette ---
BG = "#F6F3EC"        # sıcak krem zemin
SURFACE = "#FFFFFF"
BORDER = "#DCD7C9"
INK = "#1A1814"        # ana metin
MUTED = "#8A8473"      # ikincil metin
ACCENT = "#E8542C"      # mercan/turuncu vurgu
ACCENT_HOVER = "#C8431F"
ACCENT_SOFT = "#FBEAE3"
SUCCESS = "#1F6E43"
SUCCESS_SOFT = "#E7F2EA"
ERROR = "#B23A2E"
ERROR_SOFT = "#FBEAE6"

FONT_HEAD = ("Georgia", 22, "bold")
FONT_LABEL = ("Segoe UI", 11)
FONT_LABEL_BOLD = ("Segoe UI Semibold", 11)
FONT_BODY = ("Segoe UI", 12)
FONT_BUTTON = ("Segoe UI Semibold", 12)
FONT_MONO = ("Consolas", 10)

ctk.set_appearance_mode("light")

_FILETYPES = [
    ("Desteklenen dosyalar", " ".join(f"*.{ext}" for ext in sorted(security.SUPPORTED_EXTENSIONS))),
    ("Tüm dosyalar", "*.*"),
]


def _spaced(text: str) -> str:
    """Etiket metnine harf aralığı izlenimi vermek için karakterler arasına
    ince boşluk ekler (özel font tracking yerine ucuz ama etkili bir hile)."""
    return " ".join(text.upper())


class App(ctk.CTk, TkinterDnD.DnDWrapper):
    """customtkinter penceresine tkinterdnd2 sürükle-bırak desteği eklemek için
    ikisinin birleştirildiği kök pencere sınıfı."""

    def __init__(self) -> None:
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        self.title("Belge Dönüştürücü")
        self.geometry("560x620")
        self.minsize(520, 580)
        self.configure(fg_color=BG)

        icon_path = resource_path("assets/icon.ico")
        if icon_path.exists():
            try:
                self.iconbitmap(default=str(icon_path))
            except Exception:
                pass  # ikon yüklenemese de uygulama açılmaya devam etmeli

        self.source_path: Path | None = None
        self.source_ext: str | None = None
        self.output_dir: Path | None = None
        self.target_ext: str | None = None
        self.result_queue: "queue.Queue[tuple]" = queue.Queue()
        self._target_buttons: dict[str, ctk.CTkButton] = {}

        self._build_widgets()

    def report_callback_exception(self, exc, val, tb) -> None:
        """Tk olay callback'lerindeki istisnalar sys.excepthook'tan GEÇMEZ
        (Tk kendi içinde yakalar); main.py'deki crash logger'a yönlendir."""
        logging.error("Tk callback istisnası:\n%s", "".join(traceback.format_exception(exc, val, tb)))

    # ---------------------------------------------------------------- UI ---
    def _build_widgets(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        # İnce üst vurgu şeridi — marka dokunuşu, tipik AI arayüzlerinde
        # görülmeyen bir detay.
        top_bar = ctk.CTkFrame(self, height=5, corner_radius=0, fg_color=ACCENT)
        top_bar.grid(row=0, column=0, sticky="ew")

        header = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        header.grid(row=1, column=0, padx=28, pady=(26, 6), sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Belge Dönüştürücü", font=FONT_HEAD, text_color=INK).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkLabel(
            header,
            text="docx · xlsx · pptx · pdf · jpg · png",
            font=FONT_LABEL,
            text_color=MUTED,
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        body = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        body.grid(row=2, column=0, padx=28, pady=10, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- Sürükle-bırak alanı: kesik çizgili kenarlık Canvas ile çizilir ---
        self.drop_canvas = ctk.CTkCanvas(
            body, height=150, bg=SURFACE, highlightthickness=0
        )
        self.drop_canvas.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        self.drop_canvas.bind("<Configure>", self._draw_drop_border)
        self.drop_canvas.bind("<Button-1>", self._on_browse_click)
        self.drop_canvas.drop_target_register(DND_FILES)
        self.drop_canvas.dnd_bind("<<Drop>>", self._on_drop)

        self.drop_text_id = None  # canvas üzerine çizilen metin öğesi

        self.file_chip = ctk.CTkFrame(body, corner_radius=0, fg_color=SURFACE, border_width=1, border_color=BORDER)
        # ilk açılışta gizli; bir dosya seçilince grid'e eklenir
        self.file_chip_label = ctk.CTkLabel(
            self.file_chip, text="", font=FONT_MONO, text_color=INK, anchor="w"
        )
        self.file_chip_label.pack(side="left", padx=12, pady=8, fill="x", expand=True)

        # --- Hedef format: segmented düz butonlar ---
        section = ctk.CTkLabel(body, text=_spaced("Hedef format"), font=FONT_LABEL_BOLD, text_color=MUTED)
        section.grid(row=2, column=0, sticky="w", pady=(18, 6))

        self.target_row = ctk.CTkFrame(body, corner_radius=0, fg_color="transparent")
        self.target_row.grid(row=3, column=0, sticky="ew")
        self.target_placeholder = ctk.CTkLabel(
            self.target_row, text="Önce bir dosya seçin", font=FONT_LABEL, text_color=MUTED
        )
        self.target_placeholder.pack(side="left", pady=2)

        # --- Çıktı klasörü ---
        out_section = ctk.CTkLabel(body, text=_spaced("Çıktı klasörü"), font=FONT_LABEL_BOLD, text_color=MUTED)
        out_section.grid(row=4, column=0, sticky="w", pady=(20, 6))

        output_row = ctk.CTkFrame(body, corner_radius=0, fg_color="transparent")
        output_row.grid(row=5, column=0, sticky="ew")
        output_row.grid_columnconfigure(0, weight=1)
        self.output_dir_label = ctk.CTkLabel(
            output_row, text="—", font=FONT_MONO, text_color=INK, anchor="w"
        )
        self.output_dir_label.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(
            output_row,
            text="değiştir",
            width=70,
            height=26,
            corner_radius=0,
            fg_color="transparent",
            hover_color=ACCENT_SOFT,
            text_color=ACCENT,
            border_width=1,
            border_color=ACCENT,
            font=FONT_LABEL,
            command=self._choose_output_dir,
        ).grid(row=0, column=1, padx=(10, 0))

        # --- Dönüştür butonu ---
        self.convert_button = ctk.CTkButton(
            body,
            text=_spaced("Dönüştür"),
            height=46,
            corner_radius=0,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="#FFFFFF",
            font=FONT_BUTTON,
            state="disabled",
            command=self._on_convert_click,
        )
        self.convert_button.grid(row=6, column=0, sticky="ew", pady=(24, 0))

        self.progress_bar = ctk.CTkProgressBar(
            body, mode="indeterminate", height=3, corner_radius=0,
            fg_color=BORDER, progress_color=ACCENT,
        )

        # --- Durum kartı ---
        self.status_card = ctk.CTkFrame(
            body, corner_radius=0, fg_color=SURFACE, border_width=0,
        )
        self.status_dot = ctk.CTkFrame(self.status_card, width=4, fg_color=ACCENT, corner_radius=0)
        self.status_dot.pack(side="left", fill="y")
        self.status_label = ctk.CTkLabel(
            self.status_card, text="", font=FONT_BODY, text_color=INK,
            wraplength=440, justify="left", anchor="w",
        )
        self.status_label.pack(side="left", padx=14, pady=12, fill="both", expand=True)

        self.open_folder_button = ctk.CTkButton(
            body,
            text=_spaced("Klasörü aç"),
            height=36,
            corner_radius=0,
            fg_color="transparent",
            hover_color=ACCENT_SOFT,
            text_color=INK,
            border_width=1,
            border_color=BORDER,
            font=FONT_LABEL_BOLD,
            command=self._open_output_folder,
        )

    def _draw_drop_border(self, _event=None) -> None:
        c = self.drop_canvas
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 2 or h < 2:
            return
        c.create_rectangle(2, 2, w - 2, h - 2, outline=BORDER, width=2, dash=(6, 4))
        c.create_text(
            w // 2, h // 2,
            text="dosyayı buraya sürükle  ·  ya da tıkla",
            fill=MUTED, font=FONT_BODY,
        )

    # ----------------------------------------------------------- helpers ---
    def _on_browse_click(self, _event=None) -> None:
        path = ctk.filedialog.askopenfilename(title="Dosya seç", filetypes=_FILETYPES)
        if path:
            self._set_source(Path(path))

    def _on_drop(self, event) -> None:
        paths = self.tk.splitlist(event.data)
        if paths:
            self._set_source(Path(paths[0]))

    def _clear_status(self) -> None:
        self.status_card.grid_forget()
        self.open_folder_button.grid_forget()

    def _set_source(self, path: Path) -> None:
        self._clear_status()

        actual_ext = controller.detect_source_format(path)
        if actual_ext is None:
            self.source_path = None
            self.source_ext = None
            self._show_file_chip(f"⚠ tanınmayan dosya: {path.name}", error=True)
            self._set_targets([])
            return

        targets = controller.available_targets(actual_ext)
        if not targets:
            self.source_path = None
            self.source_ext = None
            self._show_file_chip(f"⚠ .{actual_ext} için desteklenen dönüşüm yok", error=True)
            self._set_targets([])
            return

        self.source_path = path
        self.source_ext = actual_ext
        self.output_dir = path.parent
        self.output_dir_label.configure(text=str(self.output_dir))

        suspicious = path.suffix.lstrip(".").lower() not in (
            actual_ext, "jpeg" if actual_ext == "jpg" else actual_ext
        )
        note = "  (uzantı içerikle uyuşmuyor)" if suspicious else ""
        self._show_file_chip(f"{path.name}   [.{actual_ext}]{note}", error=False)

        self._set_targets(targets)
        self.convert_button.configure(state="normal")

    def _show_file_chip(self, text: str, error: bool) -> None:
        self.file_chip_label.configure(text=text, text_color=ERROR if error else INK)
        self.file_chip.grid(row=1, column=0, sticky="ew", pady=(8, 0))

    def _set_targets(self, targets: list[str]) -> None:
        for child in self.target_row.winfo_children():
            child.destroy()
        self._target_buttons.clear()

        if not targets:
            self.target_placeholder = ctk.CTkLabel(
                self.target_row, text="Bu dosya için desteklenen dönüşüm yok", font=FONT_LABEL, text_color=MUTED
            )
            self.target_placeholder.pack(side="left", pady=2)
            self.target_ext = None
            self.convert_button.configure(state="disabled")
            return

        for i, ext in enumerate(targets):
            btn = ctk.CTkButton(
                self.target_row,
                text=f".{ext}",
                width=64,
                height=32,
                corner_radius=0,
                font=FONT_LABEL_BOLD,
                command=lambda e=ext: self._select_target(e),
            )
            btn.pack(side="left", padx=(0 if i == 0 else 8, 0))
            self._target_buttons[ext] = btn

        self._select_target(targets[0])

    def _select_target(self, ext: str) -> None:
        self.target_ext = ext
        for e, btn in self._target_buttons.items():
            if e == ext:
                btn.configure(fg_color=INK, text_color="#FFFFFF", hover_color=INK)
            else:
                btn.configure(fg_color="transparent", text_color=INK, hover_color=ACCENT_SOFT)
                btn.configure(border_width=1, border_color=BORDER)

    def _choose_output_dir(self) -> None:
        chosen = ctk.filedialog.askdirectory(title="Çıktı klasörü seç")
        if chosen:
            self.output_dir = Path(chosen)
            self.output_dir_label.configure(text=str(self.output_dir))

    def _open_output_folder(self) -> None:
        if self.output_dir is not None:
            os.startfile(self.output_dir)  # noqa: S606 - yalnızca Windows hedefleniyor

    # ------------------------------------------------------------ convert ---
    def _on_convert_click(self) -> None:
        if self.source_path is None or self.target_ext is None or self.output_dir is None:
            return

        self.convert_button.configure(state="disabled")
        for btn in self._target_buttons.values():
            btn.configure(state="disabled")
        self._clear_status()

        self.progress_bar.grid(row=7, column=0, sticky="ew", pady=(16, 0))
        self.progress_bar.start()

        job = ConversionJob(source=self.source_path, target_format=self.target_ext, output_dir=self.output_dir)
        self.result_queue = queue.Queue()
        controller.run_conversion_async(job, self.result_queue)
        self.after(100, self._poll_queue)

    def _poll_queue(self) -> None:
        try:
            kind, payload = self.result_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_queue)
            return

        if kind == "started":
            self.after(100, self._poll_queue)
            return

        self.progress_bar.stop()
        self.progress_bar.grid_forget()
        self.convert_button.configure(state="normal")
        for btn in self._target_buttons.values():
            btn.configure(state="normal")

        if kind == "done":
            self.status_dot.configure(fg_color=SUCCESS)
            self.status_card.configure(fg_color=SUCCESS_SOFT)
            self.status_label.configure(text=payload.message, text_color=SUCCESS)
            self.status_card.grid(row=8, column=0, sticky="ew", pady=(14, 0))
            self.open_folder_button.grid(row=9, column=0, sticky="ew", pady=(10, 0))
        else:
            self.status_dot.configure(fg_color=ERROR)
            self.status_card.configure(fg_color=ERROR_SOFT)
            self.status_label.configure(text=payload, text_color=ERROR)
            self.status_card.grid(row=8, column=0, sticky="ew", pady=(14, 0))


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
