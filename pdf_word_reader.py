
import json
import os
import re
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PyPDF2 import PdfReader

try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except Exception:
    fitz = None
    HAS_PYMUPDF = False

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except Exception:
    Image = None
    ImageTk = None
    HAS_PIL = False

CONFIG_FILE = "pdf_word_reader_config.json"


class PDFWordReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF One-Word Reader")
        self.root.geometry("1440x860")
        self.root.minsize(1100, 650)

        self.page_tokens = []
        self.play_tokens = []
        self.total_pages = 0
        self.current_index = 0  # next token index to display during autoplay

        self.playing = False
        self.paused = False
        self.in_countdown = False
        self.countdown_remaining = 0
        self.after_id = None
        self.countdown_after_id = None

        self.current_pdf_path = ""
        self.current_preview_photo = None
        self.current_doc = None
        self.preview_page_number = None

        self.session_start_time = None
        self.session_displayed_words = 0

        self.preferences = self.load_preferences()

        self.build_ui()
        self.apply_preferences_to_ui()
        self.apply_theme()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<F11>", self.toggle_fullscreen_event)
        self.root.bind("<Escape>", self.exit_fullscreen_event)
        self.root.bind("<space>", self.space_toggle_event)
        self.root.bind("<Left>", self.prev_word_event)
        self.root.bind("<Right>", self.next_word_event)
        self.page_listbox.bind("<<ListboxSelect>>", self.on_page_list_select)
        self.preview_label.bind("<Button-1>", self.on_preview_click)
        self.preview_frame.bind("<Button-1>", self.on_preview_click)
        self.preview_label.bind("<Double-Button-1>", self.on_preview_double_click)
        self.preview_frame.bind("<Double-Button-1>", self.on_preview_double_click)

        self.root.after(100, self.auto_restore_session)

    # -----------------------------
    # UI
    # -----------------------------
    def build_ui(self):
        self.main_container = tk.Frame(self.root)
        self.main_container.pack(fill="both", expand=True)

        self.top_frame = tk.Frame(self.main_container, padx=10, pady=10)
        self.top_frame.pack(fill="x")

        tk.Label(self.top_frame, text="PDF Location:").grid(row=0, column=0, sticky="w")

        self.pdf_path_var = tk.StringVar()
        self.pdf_entry = tk.Entry(self.top_frame, textvariable=self.pdf_path_var, width=90)
        self.pdf_entry.grid(row=0, column=1, padx=5, sticky="ew")

        self.browse_btn = tk.Button(self.top_frame, text="Browse", command=self.browse_pdf, width=12)
        self.browse_btn.grid(row=0, column=2, padx=5)

        self.load_btn = tk.Button(self.top_frame, text="Load PDF", command=self.load_pdf, width=12)
        self.load_btn.grid(row=0, column=3, padx=5)

        self.resume_session_btn = tk.Button(
            self.top_frame, text="Resume Last Session", command=self.resume_last_session, width=18
        )
        self.resume_session_btn.grid(row=0, column=4, padx=5)

        self.top_frame.columnconfigure(1, weight=1)

        self.settings_frame = tk.Frame(self.main_container, padx=10, pady=5)
        self.settings_frame.pack(fill="x")

        tk.Label(self.settings_frame, text="Word Interval (ms):").grid(row=0, column=0, sticky="w")
        self.interval_scale = tk.Scale(
            self.settings_frame,
            from_=50,
            to=2000,
            orient="horizontal",
            length=240,
            resolution=10,
            command=self.on_interval_change,
        )
        self.interval_scale.grid(row=0, column=1, padx=5, sticky="w")

        self.interval_value_label = tk.Label(self.settings_frame, text="")
        self.interval_value_label.grid(row=0, column=2, padx=(0, 20), sticky="w")

        tk.Label(self.settings_frame, text="Countdown (sec):").grid(row=0, column=3, sticky="w")
        self.countdown_scale = tk.Scale(
            self.settings_frame,
            from_=0,
            to=10,
            orient="horizontal",
            length=200,
            resolution=1,
            command=self.on_countdown_change,
        )
        self.countdown_scale.grid(row=0, column=4, padx=5, sticky="w")

        self.countdown_value_label = tk.Label(self.settings_frame, text="")
        self.countdown_value_label.grid(row=0, column=5, padx=(0, 20), sticky="w")

        tk.Label(self.settings_frame, text="Font Size:").grid(row=1, column=0, sticky="w")
        self.font_size_scale = tk.Scale(
            self.settings_frame,
            from_=18,
            to=140,
            orient="horizontal",
            length=240,
            resolution=1,
            command=self.on_font_size_change,
        )
        self.font_size_scale.grid(row=1, column=1, padx=5, sticky="w")

        self.font_size_value_label = tk.Label(self.settings_frame, text="")
        self.font_size_value_label.grid(row=1, column=2, padx=(0, 20), sticky="w")

        tk.Label(self.settings_frame, text="Start From Page:").grid(row=1, column=3, sticky="w")
        self.start_page_var = tk.StringVar(value="1")
        self.start_page_spinbox = tk.Spinbox(
            self.settings_frame,
            from_=1,
            to=1,
            width=10,
            textvariable=self.start_page_var,
        )
        self.start_page_spinbox.grid(row=1, column=4, padx=5, sticky="w")

        self.page_info_var = tk.StringVar(value="Pages: 0")
        self.page_info_label = tk.Label(self.settings_frame, textvariable=self.page_info_var)
        self.page_info_label.grid(row=1, column=5, padx=(0, 20), sticky="w")

        self.toggle_frame = tk.Frame(self.main_container, padx=10, pady=8)
        self.toggle_frame.pack(fill="x")

        self.dark_mode_var = tk.BooleanVar(value=False)
        self.fullscreen_var = tk.BooleanVar(value=False)

        self.dark_mode_check = tk.Checkbutton(
            self.toggle_frame, text="Dark Mode", variable=self.dark_mode_var, command=self.on_dark_mode_toggle
        )
        self.dark_mode_check.pack(side="left", padx=5)

        self.fullscreen_check = tk.Checkbutton(
            self.toggle_frame,
            text="Fullscreen",
            variable=self.fullscreen_var,
            command=self.on_fullscreen_toggle,
        )
        self.fullscreen_check.pack(side="left", padx=5)

        self.fullscreen_hint = tk.Label(
            self.toggle_frame,
            text="F11: Fullscreen | Esc: Exit | Space: Play/Pause | Left/Right: Prev/Next",
        )
        self.fullscreen_hint.pack(side="left", padx=15)

        self.button_frame = tk.Frame(self.main_container, padx=10, pady=10)
        self.button_frame.pack(fill="x")

        self.play_btn = tk.Button(self.button_frame, text="Play", command=self.play, width=12)
        self.play_btn.pack(side="left", padx=5)

        self.pause_btn = tk.Button(self.button_frame, text="Pause", command=self.pause, width=12)
        self.pause_btn.pack(side="left", padx=5)

        self.resume_btn = tk.Button(self.button_frame, text="Resume", command=self.resume, width=12)
        self.resume_btn.pack(side="left", padx=5)

        self.stop_btn = tk.Button(self.button_frame, text="Stop", command=self.stop, width=12)
        self.stop_btn.pack(side="left", padx=5)

        self.prev_btn = tk.Button(self.button_frame, text="Previous Word", command=self.previous_word, width=14)
        self.prev_btn.pack(side="left", padx=5)

        self.next_btn = tk.Button(self.button_frame, text="Next Word", command=self.next_word, width=14)
        self.next_btn.pack(side="left", padx=5)

        self.info_frame = tk.Frame(self.main_container, padx=10, pady=5)
        self.info_frame.pack(fill="x")

        self.status_var = tk.StringVar(value="No PDF loaded.")
        self.status_label = tk.Label(self.info_frame, textvariable=self.status_var, anchor="w")
        self.status_label.pack(fill="x")

        self.progress_var = tk.StringVar(value="Word 0 / 0")
        self.progress_label = tk.Label(self.info_frame, textvariable=self.progress_var, anchor="w")
        self.progress_label.pack(fill="x")

        self.page_status_var = tk.StringVar(value="Current Page: -")
        self.page_status_label = tk.Label(self.info_frame, textvariable=self.page_status_var, anchor="w")
        self.page_status_label.pack(fill="x")

        self.wpm_var = tk.StringVar(value="Configured WPM: 0 | Session WPM: 0")
        self.wpm_label = tk.Label(self.info_frame, textvariable=self.wpm_var, anchor="w")
        self.wpm_label.pack(fill="x")

        self.progress_frame = tk.Frame(self.main_container, padx=10, pady=5)
        self.progress_frame.pack(fill="x")

        self.progress_bar = ttk.Progressbar(self.progress_frame, mode="determinate", maximum=100, value=0)
        self.progress_bar.pack(fill="x")

        self.content_frame = tk.Frame(self.main_container, padx=10, pady=10)
        self.content_frame.pack(fill="both", expand=True)

        self.display_container = tk.Frame(self.content_frame)
        self.display_container.pack(side="left", fill="both", expand=True)

        self.display_frame = tk.Frame(self.display_container, bd=2, relief="sunken")
        self.display_frame.pack(fill="both", expand=True)

        self.word_label = tk.Label(
            self.display_frame,
            text="Load a PDF to begin",
            font=("Arial", 36, "bold"),
            wraplength=1000,
            justify="center",
        )
        self.word_label.place(relx=0.5, rely=0.5, anchor="center")

        self.sidebar = tk.Frame(self.content_frame, width=320)
        self.sidebar.pack(side="right", fill="y", padx=(10, 0))
        self.sidebar.pack_propagate(False)

        self.page_panel_title = tk.Label(self.sidebar, text="Pages")
        self.page_panel_title.pack(anchor="w", padx=5, pady=(0, 5))

        self.page_listbox = tk.Listbox(self.sidebar, exportselection=False, height=12)
        self.page_listbox.pack(fill="x", padx=5, pady=(0, 10))

        self.preview_title = tk.Label(self.sidebar, text="Page Preview")
        self.preview_title.pack(anchor="w", padx=5, pady=(0, 5))

        self.preview_hint = tk.Label(
            self.sidebar,
            text="Single click left/right half to preview previous/next page. Double-click preview to jump reading to that page.",
            justify="left",
            wraplength=280,
            anchor="w",
        )
        self.preview_hint.pack(anchor="w", padx=5, pady=(0, 8), fill="x")

        self.preview_frame = tk.Frame(self.sidebar, bd=1, relief="sunken")
        self.preview_frame.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        self.preview_label = tk.Label(self.preview_frame, text="Preview unavailable")
        self.preview_label.pack(fill="both", expand=True)

    # -----------------------------
    # Preferences / session persistence
    # -----------------------------
    def load_preferences(self):
        defaults = {
            "countdown": 3,
            "interval_ms": 300,
            "font_size": 36,
            "dark_mode": False,
            "fullscreen": False,
            "last_pdf_path": "",
            "last_start_page": 1,
            "last_current_index": 0,
            "last_total_play_tokens": 0,
        }
        if not os.path.exists(CONFIG_FILE):
            return defaults
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for key, value in defaults.items():
                if key not in data:
                    data[key] = value
            return data
        except Exception:
            return defaults

    def save_preferences(self):
        self.preferences["countdown"] = int(self.countdown_scale.get())
        self.preferences["interval_ms"] = int(self.interval_scale.get())
        self.preferences["font_size"] = int(self.font_size_scale.get())
        self.preferences["dark_mode"] = bool(self.dark_mode_var.get())
        self.preferences["fullscreen"] = bool(self.fullscreen_var.get())
        self.preferences["last_pdf_path"] = self.current_pdf_path or self.pdf_path_var.get().strip()
        self.preferences["last_start_page"] = self.safe_get_start_page()
        self.preferences["last_current_index"] = int(self.current_index)
        self.preferences["last_total_play_tokens"] = len(self.play_tokens)
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.preferences, f, indent=2)
        except Exception:
            pass

    def apply_preferences_to_ui(self):
        self.interval_scale.set(self.preferences.get("interval_ms", 300))
        self.countdown_scale.set(self.preferences.get("countdown", 3))
        self.font_size_scale.set(self.preferences.get("font_size", 36))
        self.dark_mode_var.set(self.preferences.get("dark_mode", False))
        self.fullscreen_var.set(self.preferences.get("fullscreen", False))
        last_path = self.preferences.get("last_pdf_path", "")
        if last_path:
            self.pdf_path_var.set(last_path)

        self.update_interval_label()
        self.update_countdown_label()
        self.update_font_size_label()
        self.update_word_font()
        self.update_wpm_labels()

        self.root.after(50, lambda: self.root.attributes("-fullscreen", self.fullscreen_var.get()))

    def auto_restore_session(self):
        last_path = self.preferences.get("last_pdf_path", "")
        if not last_path or not os.path.exists(last_path):
            return
        start_page = self.preferences.get("last_start_page", 1)
        self.start_page_var.set(str(start_page))
        self.load_pdf(silent=True)
        self.restore_last_position(show_status=True)

    def resume_last_session(self):
        last_path = self.preferences.get("last_pdf_path", "")
        if not last_path or not os.path.exists(last_path):
            messagebox.showwarning("Warning", "No resumable session found.")
            return
        self.pdf_path_var.set(last_path)
        self.start_page_var.set(str(self.preferences.get("last_start_page", 1)))
        self.load_pdf(silent=True)
        restored = self.restore_last_position(show_status=True)
        if restored:
            self.resume()

    def restore_last_position(self, show_status=False):
        if not self.page_tokens:
            return False

        self.play_tokens = self.prepare_tokens_from_start_page()
        if not self.play_tokens:
            return False

        saved_index = int(self.preferences.get("last_current_index", 0))
        if saved_index < 0:
            saved_index = 0
        if saved_index > len(self.play_tokens):
            saved_index = len(self.play_tokens)

        self.current_index = saved_index
        self.playing = False
        self.paused = False
        self.in_countdown = False

        if self.current_index > 0:
            shown_index = self.current_index - 1
            self.display_token_at(shown_index, count_for_wpm=False, persist=False)
        else:
            start_page = self.safe_get_start_page()
            self.word_label.config(text="Session restored. Press Resume or Play.")
            self.page_status_var.set(f"Current Page: {start_page}")
            self.highlight_page(start_page)
            self.render_page_preview(start_page)
            self.progress_var.set(f"Word 0 / {len(self.play_tokens)}")
            self.set_progress(0, len(self.play_tokens))

        self.session_start_time = None
        self.session_displayed_words = 0
        self.update_wpm_labels()
        self.save_preferences()

        if show_status:
            self.status_var.set(
                f"Session restored from word {self.current_index} of {len(self.play_tokens)}. Press Resume to continue."
            )
        return True

    # -----------------------------
    # File loading
    # -----------------------------
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.pdf_path_var.set(file_path)

    def close_current_doc(self):
        if self.current_doc is not None:
            try:
                self.current_doc.close()
            except Exception:
                pass
            self.current_doc = None

    def load_pdf(self, silent=False):
        path = self.pdf_path_var.get().strip()
        if not path:
            if not silent:
                messagebox.showerror("Error", "Enter a PDF path or browse for a file.")
            return False

        if not os.path.exists(path):
            if not silent:
                messagebox.showerror("Error", "PDF path does not exist.")
            return False

        try:
            reader = PdfReader(path)
            self.page_tokens = []
            self.total_pages = len(reader.pages)
            self.current_pdf_path = path
            self.preview_page_number = None

            if self.total_pages == 0:
                if not silent:
                    messagebox.showerror("Error", "The PDF appears to have no pages.")
                return False

            total_word_count = 0
            for i, page in enumerate(reader.pages, start=1):
                page_text = page.extract_text() or ""
                words = self.tokenize_text(page_text)
                page_entries = [{"word": w, "page": i} for w in words]
                self.page_tokens.append(page_entries)
                total_word_count += len(page_entries)

            if total_word_count == 0:
                if not silent:
                    messagebox.showerror(
                        "Error",
                        "No extractable text found in the PDF.\nIf this is a scanned PDF, OCR is required.",
                    )
                return False

            self.start_page_spinbox.config(to=self.total_pages)
            if self.safe_get_start_page() > self.total_pages:
                self.start_page_var.set("1")

            self.play_tokens = []
            self.current_index = 0
            self.playing = False
            self.paused = False
            self.in_countdown = False
            self.countdown_remaining = 0
            self.cancel_scheduled_tasks()
            self.refresh_page_listbox()
            self.page_info_var.set(f"Pages: {self.total_pages}")
            self.status_var.set(
                f"Loaded PDF successfully. Total pages: {self.total_pages} | Total readable words: {total_word_count}"
            )
            self.progress_var.set("Word 0 / 0")
            self.page_status_var.set("Current Page: -")
            self.set_progress(0, 0)
            self.session_start_time = None
            self.session_displayed_words = 0
            self.update_wpm_labels()

            start_page = self.safe_get_start_page()
            self.highlight_page(start_page)
            self.word_label.config(text="PDF loaded. Press Play.")
            self.open_preview_document()
            self.render_page_preview(start_page)
            self.save_preferences()
            return True
        except Exception as e:
            if not silent:
                messagebox.showerror("Error", f"Failed to load PDF:\n{e}")
            return False

    def open_preview_document(self):
        self.close_current_doc()
        if not HAS_PYMUPDF:
            self.preview_label.config(text="Preview unavailable. Install PyMuPDF.", image="")
            self.current_preview_photo = None
            return
        try:
            self.current_doc = fitz.open(self.current_pdf_path)
        except Exception:
            self.current_doc = None
            self.preview_label.config(text="Preview failed to open.", image="")
            self.current_preview_photo = None

    def tokenize_text(self, text):
        raw_tokens = re.findall(r"\S+", text)
        return [tok for tok in raw_tokens if self.is_not_punctuation_only(tok)]

    def is_not_punctuation_only(self, token):
        return any(ch.isalnum() for ch in token)

    def safe_get_start_page(self):
        try:
            page = int(self.start_page_var.get())
        except Exception:
            page = 1
        if self.total_pages <= 0:
            return 1
        if page < 1:
            page = 1
        if page > self.total_pages:
            page = self.total_pages
        return page

    def prepare_tokens_from_start_page(self):
        if not self.page_tokens:
            return []
        start_page = self.safe_get_start_page()
        flat = []
        for page_entries in self.page_tokens[start_page - 1 :]:
            flat.extend(page_entries)
        return flat

    # -----------------------------
    # Playback
    # -----------------------------
    def play(self):
        if not self.page_tokens:
            messagebox.showwarning("Warning", "Load a PDF first.")
            return

        self.cancel_scheduled_tasks()
        self.playing = False
        self.paused = False
        self.in_countdown = False
        self.current_index = 0

        self.play_tokens = self.prepare_tokens_from_start_page()
        start_page = self.safe_get_start_page()
        if not self.play_tokens:
            messagebox.showwarning("Warning", f"No readable words found from page {start_page} onward.")
            return

        self.progress_var.set(f"Word 0 / {len(self.play_tokens)}")
        self.status_var.set(f"Starting from page {start_page}. Total words in playback: {len(self.play_tokens)}")
        self.page_status_var.set(f"Current Page: {start_page}")
        self.highlight_page(start_page)
        self.render_page_preview(start_page)
        self.set_progress(0, len(self.play_tokens))

        self.session_start_time = None
        self.session_displayed_words = 0
        self.update_wpm_labels()

        self.countdown_remaining = int(self.countdown_scale.get())
        self.save_preferences()

        if self.countdown_remaining <= 0:
            self.start_reading()
        else:
            self.in_countdown = True
            self.run_countdown()

    def run_countdown(self):
        self.playing = False
        if self.countdown_remaining > 3:
            self.word_label.config(text=str(self.countdown_remaining))
        elif self.countdown_remaining == 3:
            self.word_label.config(text="Ready")
        elif self.countdown_remaining == 2:
            self.word_label.config(text="Set")
        elif self.countdown_remaining == 1:
            self.word_label.config(text="Go")
        else:
            self.in_countdown = False
            self.start_reading()
            return
        self.countdown_after_id = self.root.after(1000, self.countdown_tick)

    def countdown_tick(self):
        self.countdown_remaining -= 1
        self.run_countdown()

    def start_reading(self):
        self.playing = True
        self.paused = False
        self.in_countdown = False
        if self.session_start_time is None:
            self.session_start_time = time.time()
        self.show_next_word()

    def show_next_word(self):
        if not self.playing or self.paused:
            return

        if self.current_index >= len(self.play_tokens):
            self.word_label.config(text="Done")
            self.status_var.set("Finished reading all words.")
            self.progress_var.set(f"Word {len(self.play_tokens)} / {len(self.play_tokens)}")
            self.set_progress(len(self.play_tokens), len(self.play_tokens))
            self.playing = False
            self.after_id = None
            self.save_preferences()
            self.update_wpm_labels()
            return

        self.display_token_at(self.current_index, count_for_wpm=True, persist=False)
        self.current_index += 1
        self.save_preferences()

        interval = int(self.interval_scale.get())
        self.after_id = self.root.after(interval, self.show_next_word)

    def display_token_at(self, index, count_for_wpm=False, persist=True):
        if not self.play_tokens:
            self.word_label.config(text="No words")
            self.page_status_var.set("Current Page: -")
            self.progress_var.set("Word 0 / 0")
            self.set_progress(0, 0)
            self.preview_label.config(text="Preview unavailable", image="")
            self.current_preview_photo = None
            return

        if index < 0:
            index = 0
        if index >= len(self.play_tokens):
            index = len(self.play_tokens) - 1

        token = self.play_tokens[index]
        word = token["word"]
        page = token["page"]

        self.word_label.config(text=word)
        self.progress_var.set(f"Word {index + 1} / {len(self.play_tokens)}")
        self.page_status_var.set(f"Current Page: {page}")
        self.set_progress(index + 1, len(self.play_tokens))
        self.highlight_page(page)
        self.render_page_preview(page)

        if count_for_wpm:
            if self.session_start_time is None:
                self.session_start_time = time.time()
            self.session_displayed_words += 1
        self.update_wpm_labels()

        if persist:
            self.save_preferences()

    def pause(self):
        if self.in_countdown or (self.playing and not self.paused):
            self.paused = True
            self.cancel_scheduled_tasks()
            if self.in_countdown:
                self.status_var.set(f"Countdown paused at {self.countdown_remaining} sec.")
            else:
                self.status_var.set("Paused.")
            self.save_preferences()

    def resume(self):
        if not self.play_tokens and not self.in_countdown:
            if self.page_tokens:
                self.play_tokens = self.prepare_tokens_from_start_page()
            else:
                return

        if self.paused:
            self.paused = False
            if self.in_countdown:
                self.status_var.set("Countdown resumed.")
                self.run_countdown()
                return
            self.playing = True
            if self.session_start_time is None:
                self.session_start_time = time.time()
            self.status_var.set("Resumed.")
            self.show_next_word()
            return

        if not self.playing and self.play_tokens:
            self.paused = False
            self.playing = True
            if self.session_start_time is None:
                self.session_start_time = time.time()
            self.status_var.set("Resumed.")
            self.show_next_word()

    def stop(self):
        self.cancel_scheduled_tasks()
        self.playing = False
        self.paused = False
        self.in_countdown = False
        self.countdown_remaining = 0
        self.current_index = 0
        self.session_start_time = None
        self.session_displayed_words = 0

        if self.play_tokens:
            self.word_label.config(text="Stopped")
            self.progress_var.set(f"Word 0 / {len(self.play_tokens)}")
            start_page = self.safe_get_start_page()
            self.page_status_var.set(f"Current Page: {start_page}")
            self.highlight_page(start_page)
            self.render_page_preview(start_page)
            self.set_progress(0, len(self.play_tokens))
        else:
            self.word_label.config(text="Load a PDF to begin")
            self.progress_var.set("Word 0 / 0")
            self.page_status_var.set("Current Page: -")
            self.set_progress(0, 0)

        self.update_wpm_labels()
        self.status_var.set("Stopped.")
        self.save_preferences()

    def previous_word(self):
        if self.in_countdown:
            return
        if not self.play_tokens:
            self.play_tokens = self.prepare_tokens_from_start_page()
            if not self.play_tokens:
                return

        self.cancel_scheduled_tasks()
        self.playing = False
        self.paused = False

        shown_index = self.get_current_display_index()
        if shown_index <= 0:
            self.current_index = 0
            self.display_token_at(0, persist=False)
            self.status_var.set("At first word.")
            self.save_preferences()
            return

        target = shown_index - 1
        self.current_index = target + 1
        self.display_token_at(target, persist=False)
        self.status_var.set("Moved to previous word.")
        self.save_preferences()

    def next_word(self):
        if self.in_countdown:
            return
        if not self.play_tokens:
            self.play_tokens = self.prepare_tokens_from_start_page()
            if not self.play_tokens:
                return

        self.cancel_scheduled_tasks()
        self.playing = False
        self.paused = False

        shown_index = self.get_current_display_index()
        if shown_index < 0:
            target = 0
        else:
            target = shown_index + 1

        if target >= len(self.play_tokens):
            target = len(self.play_tokens) - 1
            self.current_index = len(self.play_tokens)
            self.display_token_at(target, persist=False)
            self.status_var.set("At last word.")
            self.save_preferences()
            return

        self.current_index = target + 1
        self.display_token_at(target, persist=False)
        self.status_var.set("Moved to next word.")
        self.save_preferences()

    def get_current_display_index(self):
        if not self.play_tokens:
            return -1
        idx = self.current_index - 1
        if idx < 0:
            return -1
        if idx >= len(self.play_tokens):
            return len(self.play_tokens) - 1
        return idx

    def cancel_scheduled_tasks(self):
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        if self.countdown_after_id:
            self.root.after_cancel(self.countdown_after_id)
            self.countdown_after_id = None

    # -----------------------------
    # Progress, WPM, page highlight, preview
    # -----------------------------
    def set_progress(self, current, total):
        if total <= 0:
            self.progress_bar["value"] = 0
            return
        self.progress_bar["value"] = (current / total) * 100.0

    def refresh_page_listbox(self):
        self.page_listbox.delete(0, tk.END)
        if not self.page_tokens:
            return
        for i, tokens in enumerate(self.page_tokens, start=1):
            self.page_listbox.insert(tk.END, f"Page {i}  ({len(tokens)} words)")

    def highlight_page(self, page_number):
        if self.page_listbox.size() == 0:
            return
        idx = page_number - 1
        if idx < 0 or idx >= self.page_listbox.size():
            return
        self.page_listbox.selection_clear(0, tk.END)
        self.page_listbox.selection_set(idx)
        self.page_listbox.activate(idx)
        self.page_listbox.see(idx)

    def render_page_preview(self, page_number):
        self.preview_page_number = page_number

        if not HAS_PYMUPDF or not HAS_PIL:
            msg = "Preview unavailable. Install PyMuPDF and Pillow."
            self.preview_label.config(text=msg, image="")
            self.current_preview_photo = None
            return

        if self.current_doc is None:
            self.preview_label.config(text="Preview unavailable.", image="")
            self.current_preview_photo = None
            return

        try:
            page_index = max(0, min(page_number - 1, self.current_doc.page_count - 1))
            page = self.current_doc.load_page(page_index)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False)
            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            target_w = 280
            target_h = 400
            image.thumbnail((target_w, target_h))

            photo = ImageTk.PhotoImage(image)
            self.current_preview_photo = photo
            self.preview_label.config(image=photo, text="")
        except Exception:
            self.preview_label.config(text="Preview rendering failed.", image="")
            self.current_preview_photo = None

    def preview_selected_page(self, page_number):
        if self.total_pages <= 0:
            return
        if page_number < 1:
            page_number = 1
        if page_number > self.total_pages:
            page_number = self.total_pages
        self.highlight_page(page_number)
        self.render_page_preview(page_number)
        self.page_status_var.set(f"Current Page: {page_number}")

    def jump_to_page(self, page_number, autoplay=None):
        if self.total_pages <= 0 or not self.page_tokens:
            return False

        if page_number < 1:
            page_number = 1
        if page_number > self.total_pages:
            page_number = self.total_pages

        if autoplay is None:
            autoplay = bool(self.playing and not self.paused and not self.in_countdown)

        self.cancel_scheduled_tasks()
        self.playing = False
        self.paused = False
        self.in_countdown = False
        self.countdown_remaining = 0

        self.start_page_var.set(str(page_number))
        self.play_tokens = self.prepare_tokens_from_start_page()

        if not self.play_tokens:
            self.word_label.config(text="No readable words on selected page.")
            self.page_status_var.set(f"Current Page: {page_number}")
            self.highlight_page(page_number)
            self.render_page_preview(page_number)
            self.progress_var.set("Word 0 / 0")
            self.set_progress(0, 0)
            self.current_index = 0
            self.status_var.set(f"Page {page_number} has no readable words from this starting point.")
            self.save_preferences()
            return False

        self.current_index = 0
        self.session_start_time = None
        self.session_displayed_words = 0
        self.display_token_at(0, count_for_wpm=False, persist=False)
        self.current_index = 1
        self.update_wpm_labels()
        self.status_var.set(f"Jumped to page {page_number}.")
        self.save_preferences()

        if autoplay:
            self.resume()
        return True

    def on_preview_click(self, event=None):
        if self.total_pages <= 0:
            return

        current_page = self.preview_page_number or self.safe_get_start_page()
        widget = event.widget if event is not None else self.preview_frame
        try:
            width = widget.winfo_width()
        except Exception:
            width = 1
        if width <= 0:
            width = 1
        x = getattr(event, "x", width // 2)

        if x < width / 2:
            target_page = max(1, current_page - 1)
        else:
            target_page = min(self.total_pages, current_page + 1)

        self.preview_selected_page(target_page)
        self.status_var.set(f"Previewing page {target_page}. Double-click preview to jump reading there.")

    def on_preview_double_click(self, event=None):
        if self.total_pages <= 0:
            return
        page_number = self.preview_page_number or self.safe_get_start_page()
        self.jump_to_page(page_number, autoplay=None)

    def update_wpm_labels(self):
        interval_ms = max(1, int(self.interval_scale.get()))
        configured_wpm = int(round(60000 / interval_ms))

        if self.session_start_time and self.session_displayed_words > 0:
            elapsed_minutes = max((time.time() - self.session_start_time) / 60.0, 1e-9)
            session_wpm = int(round(self.session_displayed_words / elapsed_minutes))
        else:
            session_wpm = 0

        self.wpm_var.set(f"Configured WPM: {configured_wpm} | Session WPM: {session_wpm}")

    # -----------------------------
    # Settings handlers
    # -----------------------------
    def on_interval_change(self, _value=None):
        self.update_interval_label()
        self.update_wpm_labels()
        self.save_preferences()

    def on_countdown_change(self, _value=None):
        self.update_countdown_label()
        self.save_preferences()

    def on_font_size_change(self, _value=None):
        self.update_font_size_label()
        self.update_word_font()
        self.save_preferences()

    def update_interval_label(self):
        self.interval_value_label.config(text=f"{int(self.interval_scale.get())} ms")

    def update_countdown_label(self):
        self.countdown_value_label.config(text=f"{int(self.countdown_scale.get())} sec")

    def update_font_size_label(self):
        self.font_size_value_label.config(text=f"{int(self.font_size_scale.get())} px")

    def update_word_font(self):
        size = int(self.font_size_scale.get())
        self.word_label.config(font=("Arial", size, "bold"))

    # -----------------------------
    # Theme and window
    # -----------------------------
    def on_dark_mode_toggle(self):
        self.apply_theme()
        self.save_preferences()

    def on_fullscreen_toggle(self):
        self.root.attributes("-fullscreen", self.fullscreen_var.get())
        self.save_preferences()

    def toggle_fullscreen_event(self, event=None):
        current = bool(self.fullscreen_var.get())
        self.fullscreen_var.set(not current)
        self.root.attributes("-fullscreen", self.fullscreen_var.get())
        self.save_preferences()

    def exit_fullscreen_event(self, event=None):
        self.fullscreen_var.set(False)
        self.root.attributes("-fullscreen", False)
        self.save_preferences()

    def apply_theme(self):
        dark = self.dark_mode_var.get()
        if dark:
            bg_main = "#1e1e1e"
            bg_panel = "#2b2b2b"
            bg_display = "#111111"
            fg_text = "#f5f5f5"
            entry_bg = "#333333"
            entry_fg = "#ffffff"
            word_bg = "#111111"
            word_fg = "#ffffff"
            list_bg = "#202020"
            list_fg = "#ffffff"
            select_bg = "#3a7afe"
            select_fg = "#ffffff"
        else:
            bg_main = "#f0f0f0"
            bg_panel = "#f0f0f0"
            bg_display = "#ffffff"
            fg_text = "#000000"
            entry_bg = "#ffffff"
            entry_fg = "#000000"
            word_bg = "#ffffff"
            word_fg = "#000000"
            list_bg = "#ffffff"
            list_fg = "#000000"
            select_bg = "#0a64ad"
            select_fg = "#ffffff"

        self.root.configure(bg=bg_main)

        frame_widgets = [
            self.main_container,
            self.top_frame,
            self.settings_frame,
            self.toggle_frame,
            self.button_frame,
            self.info_frame,
            self.progress_frame,
            self.content_frame,
            self.display_container,
            self.sidebar,
            self.preview_frame,
        ]
        for widget in frame_widgets:
            widget.configure(bg=bg_panel)

        label_widgets = [
            self.status_label,
            self.progress_label,
            self.page_status_label,
            self.wpm_label,
            self.interval_value_label,
            self.countdown_value_label,
            self.font_size_value_label,
            self.page_info_label,
            self.fullscreen_hint,
            self.page_panel_title,
            self.preview_title,
            self.preview_hint,
        ]
        for widget in label_widgets:
            widget.configure(bg=bg_panel, fg=fg_text)

        for child in self.top_frame.winfo_children():
            if isinstance(child, tk.Label):
                child.configure(bg=bg_panel, fg=fg_text)
        for child in self.settings_frame.winfo_children():
            if isinstance(child, tk.Label):
                child.configure(bg=bg_panel, fg=fg_text)

        self.pdf_entry.configure(bg=entry_bg, fg=entry_fg, insertbackground=entry_fg)
        self.start_page_spinbox.configure(
            bg=entry_bg, fg=entry_fg, buttonbackground=bg_panel, insertbackground=entry_fg
        )
        self.dark_mode_check.configure(
            bg=bg_panel,
            fg=fg_text,
            selectcolor=bg_panel,
            activebackground=bg_panel,
            activeforeground=fg_text,
        )
        self.fullscreen_check.configure(
            bg=bg_panel,
            fg=fg_text,
            selectcolor=bg_panel,
            activebackground=bg_panel,
            activeforeground=fg_text,
        )
        self.display_frame.configure(bg=bg_display)
        self.word_label.configure(bg=word_bg, fg=word_fg)
        self.page_listbox.configure(
            bg=list_bg,
            fg=list_fg,
            selectbackground=select_bg,
            selectforeground=select_fg,
            highlightthickness=0,
            bd=1,
        )
        self.preview_label.configure(bg=bg_display, fg=fg_text)

        style = ttk.Style()
        try:
            style.theme_use(style.theme_use())
        except Exception:
            pass
        style.configure("TProgressbar", thickness=16)

    # -----------------------------
    # Events
    # -----------------------------
    def on_page_list_select(self, event=None):
        try:
            selection = self.page_listbox.curselection()
            if not selection:
                return
            page_num = selection[0] + 1
            self.preview_selected_page(page_num)
            self.status_var.set(f"Previewing page {page_num}. Double-click preview to jump reading there.")
        except Exception:
            pass

    def space_toggle_event(self, event=None):
        if self.paused:
            self.resume()
        elif self.playing or self.in_countdown:
            self.pause()
        else:
            self.play()

    def prev_word_event(self, event=None):
        self.previous_word()

    def next_word_event(self, event=None):
        self.next_word()

    def on_close(self):
        self.save_preferences()
        self.close_current_doc()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFWordReaderApp(root)
    root.mainloop()
