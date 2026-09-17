import sys
import os
import json
import copy
import re
import subprocess
import datetime
import threading
import queue
import webbrowser
import winsound
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Menu
import tkinter.font as tkfont

# --- AUTOMATYCZNE SPRAWDZANIE I INSTALACJA BIBLIOTEK ---
def ensure_dependencies():
    packages = {
        "customtkinter": "customtkinter",
        "pypdf": "pypdf",
        "tkinterdnd2": "tkinterdnd2",
        "PyMuPDF": "pymupdf",
        "Pillow": "PIL",
        "tkcalendar": "tkcalendar"
    }
    for pkg, import_name in packages.items():
        try:
            __import__(import_name)
        except ImportError:
            print(f"[Auto-Setup] Instalowanie biblioteki: {pkg}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

ensure_dependencies()

import customtkinter as ctk
from pypdf import PdfWriter
from tkinterdnd2 import DND_FILES, TkinterDnD
import pymupdf as fitz
from PIL import Image, ImageTk, ImageDraw
from tkcalendar import Calendar

DATE_SUFFIX_RE = re.compile(r'_(\d{4}\.\d{2}\.\d{2})(\.pdf)$', re.IGNORECASE)
APP_VERSION = "0.2"
ABOUT_WEBSITE = "WWW.DRAFTCON.PL"
ABOUT_EMAIL = "biuro@draftcon.pl"
ABOUT_DATE = "09.2026"

# Standardowe formaty papieru (szer., wys. w mm, orientacja pionowa) do
# rozpoznawania rozmiaru strony PDF w oknie informacyjnym.
PAPER_SIZES_MM = [
    ("A0", 841, 1189), ("A1", 594, 841), ("A2", 420, 594),
    ("A3", 297, 420), ("A4", 210, 297), ("A5", 148, 210),
    ("Letter", 216, 279), ("Legal", 216, 356),
]

# --- KOLOROWE FLAGI RYSOWANE PROGRAMOWO (PIL) ---
# Windows w Segoe UI Emoji nie rysuje flag jako obrazków — pokazuje literowy
# kod kraju ("PL", "GB") zamiast kolorowej flagi. Rysujemy je więc same,
# w wysokiej rozdzielczości źródłowej; CTkImage przeskalowuje je w dół do
# rozmiaru docelowego z uwzględnieniem współczynnika DPI okna, więc krawędzie
# zostają gładkie na każdym monitorze.
def _draw_flag_pl(w=240, h=150):
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, h // 2, w, h], fill="#dc143c")
    draw.rectangle([0, 0, w - 1, h - 1], outline="#999999", width=2)
    return img

def _draw_flag_gb(w=240, h=150):
    img = Image.new("RGB", (w, h), "#00247d")
    draw = ImageDraw.Draw(img)

    # Szerokie białe przekątne (krzyż św. Andrzeja)
    diag_w = int(h * 0.22)
    draw.line([(0, 0), (w, h)], fill="white", width=diag_w)
    draw.line([(0, h), (w, 0)], fill="white", width=diag_w)

    # Węższe czerwone przekątne na wierzchu (przybliżenie krzyża św. Patryka)
    diag_w2 = int(h * 0.09)
    draw.line([(0, 0), (w, h)], fill="#cf142b", width=diag_w2)
    draw.line([(0, h), (w, 0)], fill="#cf142b", width=diag_w2)

    # Biały krzyż św. Jerzego (szeroki, na całą wysokość/szerokość)
    cross_w = int(h * 0.34)
    draw.rectangle([w // 2 - cross_w // 2, 0, w // 2 + cross_w // 2, h], fill="white")
    draw.rectangle([0, h // 2 - cross_w // 2, w, h // 2 + cross_w // 2], fill="white")

    # Węższy czerwony krzyż na wierzchu białego
    cross_w2 = int(h * 0.18)
    draw.rectangle([w // 2 - cross_w2 // 2, 0, w // 2 + cross_w2 // 2, h], fill="#cf142b")
    draw.rectangle([0, h // 2 - cross_w2 // 2, w, h // 2 + cross_w2 // 2], fill="#cf142b")

    draw.rectangle([0, 0, w - 1, h - 1], outline="#999999", width=2)
    return img

# --- SŁOWNIK TŁUMACZEŃ (interfejs dwujęzyczny PL/EN) ---
# Klucze wspólne dla obu języków; kod i komentarze w kodzie pozostają po
# polsku niezależnie od wybranego języka interfejsu.
TRANSLATIONS = {
    "pl": {
        "app_title": "PDF Project Tree Manager v.{ver}",
        "project_unsaved": "Projekt: (Niezapisany)",
        "project_label": "Projekt: {path}",
        "tt_open_project": "Otwórz projekt JSON",
        "tt_save_project": "Zapisz projekt",
        "tt_output_dir": "Katalog docelowy",
        "tt_new_target": "Nowy plik wynikowy",
        "tt_add_pdf": "Dodaj PDF",
        "tt_refresh": "Odśwież",
        "tt_delete": "Usuń (Del)",
        "tt_undo": "Cofnij",
        "tt_redo": "Ponów",
        "tt_lang": "Zmień język interfejsu",
        "chk_date_suffix": "Dopisuj datę",
        "preview_none": "Podgląd: (nic nie zaznaczono)",
        "preview_name": "Podgląd: {name}",
        "preview_missing": "❌ Plik nie istnieje na dysku",
        "preview_error": "⚠️ Nie można wygenerować podglądu\n(uszkodzony PDF?)",
        "btn_open_folder": "Otwórz folder",
        "btn_merge": "Scal zestawy",
        "btn_merging": "Scalanie...",
        "status_ready": " Gotowy",
        "status_hover_root": " 🏢 Wspólny katalog docelowy: {path}",
        "status_hover_target": " 📁 Docelowy plik PDF: {path}",
        "status_hover_file": " 📄 Pełna ścieżka pliku składowego: {path}",
        "warn_title": "Uwaga",
        "warn_no_output_dir": "Katalog docelowy nie jest jeszcze ustawiony lub nie istnieje.",
        "err_title": "Błąd",
        "err_file_missing": "Plik nie istnieje na dysku:\n{path}",
        "date_dialog_title": "Wybierz datę",
        "btn_select": "Wybierz",
        "btn_cancel": "Anuluj",
        "outdir_dialog_title": "Katalog docelowy",
        "outdir_dialog_label": "Ścieżka katalogu docelowego (jeśli nie istnieje, zostanie utworzona):",
        "browse_btn": "Przeglądaj...",
        "browse_dialog_title": "Przeglądaj istniejące katalogi",
        "btn_ok": "OK",
        "name_exists_warn": "Plik wynikowy o nazwie „{name}” już istnieje w projekcie.",
        "rename_dialog_title": "Nazwa pliku wynikowego",
        "root_label": "🏢 KATALOG DOCELOWY: {path}",
        "root_not_set": "❌ NIE USTAWIONO KATALOGU DOCELOWEGO",
        "info_output_dir_title": "Katalog Docelowy",
        "info_output_dir_body": "Wspólny katalog wyjściowy:\n{path}",
        "info_not_selected": "Nie wybrano",
        "page_size_error": "nie udało się odczytać",
        "orientation_h": "poziomo",
        "orientation_v": "pionowo",
        "status_missing_title": "PLIK NIE ISTNIEJE NA DYSKU",
        "status_outdated_title": "NIEAKTUALNY (zmodyfikowany po zapisaniu)",
        "status_current_title": "AKTUALNY (bez zmian)",
        "detail_saved_disk": "Data w projekcie: {saved}\nData na dysku:      {curr}",
        "detail_last_modified": "Data ostatniej modyfikacji: {curr}",
        "file_info_dialog_title": "Informacje o pliku PDF",
        "row_index": "Numer w zestawieniu:",
        "row_filename": "Nazwa pliku:",
        "row_pagesize": "Rozmiar strony:",
        "row_status": "Status:",
        "row_fullpath": "Pełna ścieżka:",
        "ctx_set_output_dir": "Ustaw katalog docelowy...",
        "ctx_open_output_folder": "Otwórz folder wynikowy",
        "ctx_new_target": "Nowy plik wynikowy...",
        "ctx_refresh": "Odśwież stany",
        "ctx_rename": "Zmień nazwę (F2)",
        "ctx_add_pdf": "Dodaj pliki PDF...",
        "ctx_details": "Szczegóły",
        "ctx_delete_target": "Usuń zestaw (Del)",
        "ctx_open_file": "Otwórz plik",
        "ctx_reveal": "Pokaż w Eksploratorze",
        "ctx_sort_az": "Sortuj (A-Z)",
        "ctx_sort_za": "Sortuj (Z-A)",
        "ctx_delete_file": "Usuń plik (Del)",
        "ctx_about": "O programie...",
        "about_title": "O programie",
        "about_app_name": "PDF Project Tree Manager",
        "about_version": "wersja {ver}",
        "about_freeware": "Freeware — program bezpłatny",
        "about_feedback": "Aby podzielić się opinią lub otrzymać nowszą wersję programu, napisz:",
        "about_version_date": "Wersja z: {date}",
        "confirm_title": "Potwierdzenie",
        "confirm_delete_targets": "Czy na pewno chcesz usunąć zaznaczone zestawy ({n}) wraz z plikami?",
        "err_no_targets": "Brak zestawów do scalenia.",
        "warn_min_target": "Projekt musi posiadać co najmniej jeden plik wynikowy.",
        "new_target_dialog_title": "Nowy Plik Wynikowy",
        "new_target_dialog_label": "Wprowadź nazwę pliku wyjściowego (np. Operat_Projektowy.pdf):",
        "merge_result_title": "Wynik scalania",
        "merge_result_text": "Zaktualizowano zestawów: {success} / {total}\n\nProblemy:\n{errors}",
        "merge_no_files": "- {name}: Brak plików składowych.",
        "success_title": "Sukces!",
        "success_text": "🔄 Pomyślnie zaktualizowano wszystkie pliki wyjściowe!",
        "open_folder_btn": "📂 Otwórz folder",
        "project_saved_title": "Projekt",
        "project_saved_text": "Zapisano projekt wraz ze wspólnym katalogiem docelowym!",
        "create_dir_err_title": "Błąd tworzenia katalogu",
        "create_dir_err_text": "Nie udało się utworzyć katalogu:\n{err}",
        "file_filter_pdf": "Pliki PDF",
        "file_filter_project": "Projekt PDF Merger",
    },
    "en": {
        "app_title": "PDF Project Tree Manager v.{ver}",
        "project_unsaved": "Project: (Unsaved)",
        "project_label": "Project: {path}",
        "tt_open_project": "Open JSON project",
        "tt_save_project": "Save project",
        "tt_output_dir": "Output directory",
        "tt_new_target": "New output file",
        "tt_add_pdf": "Add PDF",
        "tt_refresh": "Refresh",
        "tt_delete": "Delete (Del)",
        "tt_undo": "Undo",
        "tt_redo": "Redo",
        "tt_lang": "Switch interface language",
        "chk_date_suffix": "Append date",
        "preview_none": "Preview: (nothing selected)",
        "preview_name": "Preview: {name}",
        "preview_missing": "❌ File does not exist on disk",
        "preview_error": "⚠️ Unable to generate preview\n(corrupted PDF?)",
        "btn_open_folder": "Open folder",
        "btn_merge": "Merge sets",
        "btn_merging": "Merging...",
        "status_ready": " Ready",
        "status_hover_root": " 🏢 Shared output directory: {path}",
        "status_hover_target": " 📁 Target PDF file: {path}",
        "status_hover_file": " 📄 Full path of component file: {path}",
        "warn_title": "Warning",
        "warn_no_output_dir": "The output directory is not set yet or does not exist.",
        "err_title": "Error",
        "err_file_missing": "File does not exist on disk:\n{path}",
        "date_dialog_title": "Select date",
        "btn_select": "Select",
        "btn_cancel": "Cancel",
        "outdir_dialog_title": "Output directory",
        "outdir_dialog_label": "Output directory path (will be created if it doesn't exist):",
        "browse_btn": "Browse...",
        "browse_dialog_title": "Browse existing directories",
        "btn_ok": "OK",
        "name_exists_warn": "An output file named \u201c{name}\u201d already exists in the project.",
        "rename_dialog_title": "Output file name",
        "root_label": "🏢 OUTPUT DIRECTORY: {path}",
        "root_not_set": "❌ OUTPUT DIRECTORY NOT SET",
        "info_output_dir_title": "Output Directory",
        "info_output_dir_body": "Shared output directory:\n{path}",
        "info_not_selected": "Not selected",
        "page_size_error": "could not be read",
        "orientation_h": "landscape",
        "orientation_v": "portrait",
        "status_missing_title": "FILE DOES NOT EXIST ON DISK",
        "status_outdated_title": "OUTDATED (modified after saving)",
        "status_current_title": "UP TO DATE (unchanged)",
        "detail_saved_disk": "Date in project: {saved}\nDate on disk:       {curr}",
        "detail_last_modified": "Last modified: {curr}",
        "file_info_dialog_title": "PDF file information",
        "row_index": "Number in the set:",
        "row_filename": "File name:",
        "row_pagesize": "Page size:",
        "row_status": "Status:",
        "row_fullpath": "Full path:",
        "ctx_set_output_dir": "Set output directory...",
        "ctx_open_output_folder": "Open output folder",
        "ctx_new_target": "New output file...",
        "ctx_refresh": "Refresh states",
        "ctx_rename": "Rename (F2)",
        "ctx_add_pdf": "Add PDF files...",
        "ctx_details": "Details",
        "ctx_delete_target": "Delete set (Del)",
        "ctx_open_file": "Open file",
        "ctx_reveal": "Show in Explorer",
        "ctx_sort_az": "Sort (A-Z)",
        "ctx_sort_za": "Sort (Z-A)",
        "ctx_delete_file": "Delete file (Del)",
        "ctx_about": "About...",
        "about_title": "About",
        "about_app_name": "PDF Project Tree Manager",
        "about_version": "version {ver}",
        "about_freeware": "Freeware — free to use",
        "about_feedback": "To share feedback or get a newer version of the program, write to:",
        "about_version_date": "Version date: {date}",
        "confirm_title": "Confirmation",
        "confirm_delete_targets": "Are you sure you want to delete the selected sets ({n}) along with their files?",
        "err_no_targets": "No sets to merge.",
        "warn_min_target": "The project must have at least one output file.",
        "new_target_dialog_title": "New Output File",
        "new_target_dialog_label": "Enter the output file name (e.g. Project_Report.pdf):",
        "merge_result_title": "Merge result",
        "merge_result_text": "Sets updated: {success} / {total}\n\nIssues:\n{errors}",
        "merge_no_files": "- {name}: No component files.",
        "success_title": "Success!",
        "success_text": "🔄 All output files were updated successfully!",
        "open_folder_btn": "📂 Open folder",
        "project_saved_title": "Project",
        "project_saved_text": "Project saved along with the shared output directory!",
        "create_dir_err_title": "Directory creation error",
        "create_dir_err_text": "Failed to create the directory:\n{err}",
        "file_filter_pdf": "PDF files",
        "file_filter_project": "PDF Merger project",
    },
}

# --- DYMEK PODPOWIEDZI (tkinter/customtkinter nie mają go wbudowanego) ---
# Pojawia się NAD widgetem po chwili najechania myszą; pozycja liczona
# dopiero po utworzeniu etykiety, żeby dymek był naprawdę wyśrodkowany
# (stąd chwilowe wm_withdraw() zamiast migotania w złym miejscu).
class ToolTip:
    def __init__(self, widget, text, delay=500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window = None
        self._after_id = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        self._after_id = None
        if self.tip_window or not self.text:
            return

        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_withdraw()
        try:
            tw.attributes("-topmost", True)
        except Exception:
            pass

        tk.Label(
            tw, text=self.text, justify="left",
            bg="#2b2b2b", fg="white", font=("Segoe UI", 9),
            relief="solid", borderwidth=1, padx=6, pady=3
        ).pack()

        tw.update_idletasks()
        x = self.widget.winfo_rootx() + (self.widget.winfo_width() - tw.winfo_width()) // 2
        y = self.widget.winfo_rooty() - tw.winfo_height() - 6
        tw.wm_geometry(f"+{x}+{y}")
        tw.wm_deiconify()

    def _hide(self, event=None):
        self._cancel()
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

    def set_text(self, text):
        self.text = text


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PDFMultiMergerApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()

        self.TkdndVersion = TkinterDnD._require(self)

        # Język interfejsu (nie dotyczy kodu/komentarzy — te zostają po polsku)
        self.lang = "pl"

        # Ikony flag PL/GB — rysowane raz w wysokiej rozdzielczości (patrz
        # _draw_flag_pl/_draw_flag_gb); CTkImage sam przelicza docelowy
        # rozmiar wg skalowania DPI okna przy każdym renderze.
        self._flag_size = (28, 18)
        self._flag_img_pl = ctk.CTkImage(light_image=_draw_flag_pl(), dark_image=_draw_flag_pl(), size=self._flag_size)
        self._flag_img_gb = ctk.CTkImage(light_image=_draw_flag_gb(), dark_image=_draw_flag_gb(), size=self._flag_size)

        self.title(self.tr("app_title", ver=APP_VERSION))
        # customtkinter mnoży KAŻDĄ wartość przekazaną do self.geometry() przez
        # współczynnik skalowania DPI systemu (np. ×1.4 przy 140%) — bez korekty
        # "850x930" renderuje się znacznie większe. _dpi_geometry() dzieli
        # z powrotem, żeby geometry() dostawała docelowy rozmiar fizyczny.
        self.geometry(self._dpi_geometry(850, 930))  # tymczasowo, dociągnięte niżej do naturalnej szerokości przycisków

        # Wspólny katalog wyjściowy
        self.output_dir = ""

        # Data dopisywana jako końcówka nazw (PZT/PAB/... i katalogu docelowego)
        self.selected_date = datetime.date.today()
        self.date_suffix_enabled = True

        # Dynamiczne generowanie domyślnego drzewa z dzisiejszą datą (Format: YYYY.MM.DD)
        today_str = self.selected_date.strftime("%Y.%m.%d")
        self.targets = [
            {"name": f"PZT_{today_str}.pdf", "files": []},
            {"name": f"PAB_{today_str}.pdf", "files": []},
            {"name": f"ZL_{today_str}.pdf", "files": []}
        ]
        self.project_file_path = ""
        self.dragged_iids = []
        self._last_drag_target = None
        self._last_ext_drop = None
        self._editing = False
        self._merging = False
        self._autoscroll_dir = None
        self._autoscroll_job = None
        self._preview_cache = {}  # fpath -> (mtime, PIL.Image w wysokiej rozdzielczości)
        self._current_preview_path = None
        self._current_preview_ctk_img = None
        self._preview_resize_job = None
        self._undo_stack = []
        self._redo_stack = []
        self._undo_limit = 50
        self._last_moved_iids = []

        self.bind("<Control-z>", lambda e: self.undo())
        self.bind("<Control-y>", lambda e: self.redo())
        self.bind("<Control-Shift-Z>", lambda e: self.redo())

        # Przyciski jako same ikony (węższy pasek) + dymek podpowiedzi z pełną
        # nazwą akcji, bo bez tekstu ikony same w sobie są niejednoznaczne.
        # "Segoe UI" nie ma pełnego kompletu glifów emoji (część renderuje się
        # jako puste kwadraciki albo krzywo wyśrodkowana) — "Segoe UI Emoji"
        # rysuje je poprawnie na każdym monitorze/DPI. CTkFont sam przelicza
        # rozmiar wg współczynnika skalowania DPI (ScalingTracker), więc nie
        # trzeba tu ręcznie mnożyć rozmiaru.
        icon_btn_font = ctk.CTkFont(family="Segoe UI Emoji", size=15)

        # --- GÓRNY PANEL: Zarządzanie Projektem ---
        self.proj_frame = ctk.CTkFrame(self)
        self.proj_frame.pack(pady=10, padx=10, fill="x")

        # wraplength — pełna ścieżka projektu potrafi być długa; bez tego
        # rozpychałaby cały wiersz (i okno) zamiast zawinąć się do 2. linii.
        self.lbl_project = ctk.CTkLabel(
            self.proj_frame, text=self.tr("project_unsaved"),
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            wraplength=700, justify="left", anchor="w"
        )
        self.lbl_project.pack(side="left", padx=10, fill="x", expand=True)

        self.btn_lang = ctk.CTkButton(self.proj_frame, image=self._lang_flag_image(), text="", width=40, command=self.toggle_language)
        self.btn_lang.pack(side="right", padx=5, pady=5)
        self.tt_lang = ToolTip(self.btn_lang, self.tr("tt_lang"))

        self.btn_load_proj = ctk.CTkButton(self.proj_frame, text="📂", width=36, font=icon_btn_font, command=self.load_project)
        self.btn_load_proj.pack(side="right", padx=5, pady=5)
        self.tt_load_proj = ToolTip(self.btn_load_proj, self.tr("tt_open_project"))

        self.btn_save_proj = ctk.CTkButton(self.proj_frame, text="💾", width=36, font=icon_btn_font, command=self.save_project)
        self.btn_save_proj.pack(side="right", padx=5, pady=5)
        self.tt_save_proj = ToolTip(self.btn_save_proj, self.tr("tt_save_project"))

        # --- PANEL PRZYCISKÓW AKCJI ---
        self.actions_frame = ctk.CTkFrame(self)
        self.actions_frame.pack(pady=5, padx=10, fill="x")

        self.btn_katalog = ctk.CTkButton(self.actions_frame, text="🎯", width=36, font=icon_btn_font, command=self.select_output_dir)
        self.btn_katalog.pack(side="left", padx=5, pady=5)
        self.tt_katalog = ToolTip(self.btn_katalog, self.tr("tt_output_dir"))

        self.btn_new_target = ctk.CTkButton(self.actions_frame, text="🆕", width=36, font=icon_btn_font, command=self.add_target)
        self.btn_new_target.pack(side="left", padx=5, pady=5)
        self.tt_new_target = ToolTip(self.btn_new_target, self.tr("tt_new_target"))

        self.btn_add_pdf = ctk.CTkButton(self.actions_frame, text="➕", width=36, font=icon_btn_font, command=self.add_files_to_selected)
        self.btn_add_pdf.pack(side="left", padx=5, pady=5)
        self.tt_add_pdf = ToolTip(self.btn_add_pdf, self.tr("tt_add_pdf"))

        self.btn_refresh = ctk.CTkButton(self.actions_frame, text="🔄", width=36, font=icon_btn_font, command=self.refresh_tree)
        self.btn_refresh.pack(side="left", padx=5, pady=5)
        self.tt_refresh = ToolTip(self.btn_refresh, self.tr("tt_refresh"))

        self.btn_delete = ctk.CTkButton(
            self.actions_frame, text="🗑", width=36, font=icon_btn_font, anchor="center",
            fg_color="#a93226", hover_color="#7b241c", command=self.delete_selected
        )
        self.btn_delete.pack(side="left", padx=5, pady=5)
        self.tt_delete = ToolTip(self.btn_delete, self.tr("tt_delete"))

        self.btn_undo = ctk.CTkButton(self.actions_frame, text="↩", width=36, font=icon_btn_font, state="disabled", command=self.undo)
        self.btn_undo.pack(side="left", padx=(15, 5), pady=5)
        self.tt_undo = ToolTip(self.btn_undo, self.tr("tt_undo"))

        self.btn_redo = ctk.CTkButton(self.actions_frame, text="↪", width=36, font=icon_btn_font, state="disabled", command=self.redo)
        self.btn_redo.pack(side="left", padx=5, pady=5)
        self.tt_redo = ToolTip(self.btn_redo, self.tr("tt_redo"))

        # Data dopisywana jako końcówka nazw (PZT/PAB/... i katalogu docelowego)
        self.btn_date_picker = ctk.CTkButton(
            self.actions_frame,
            text=f"📅 {self.selected_date.strftime('%Y.%m.%d')}",
            width=110,
            command=self.open_date_picker
        )
        self.btn_date_picker.pack(side="right", padx=5, pady=5)

        self.date_suffix_var = tk.BooleanVar(value=self.date_suffix_enabled)
        self.chk_date_suffix = ctk.CTkCheckBox(
            self.actions_frame,
            text=self.tr("chk_date_suffix"),
            variable=self.date_suffix_var,
            command=self._toggle_date_suffix
        )
        self.chk_date_suffix.pack(side="right", padx=5, pady=5)

        # --- PODZIAŁ: DRZEWO (lewo) / PODGLĄD PDF (prawo), z przeciąganym uchwytem ---
        self.split_frame = tk.PanedWindow(self, orient="horizontal", sashwidth=6, sashrelief="raised", bd=0)
        self.split_frame.pack(pady=5, padx=10, fill="both", expand=True)

        # --- DRZEWO PROJEKTU ---
        self.tree_frame = ctk.CTkFrame(self.split_frame)

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.tree = ttk.Treeview(self.tree_frame, selectmode="extended", show="tree")
        self.tree.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)

        self.tree_scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree_scrollbar.pack(side="right", fill="y", padx=(0, 5), pady=5)
        self.tree.configure(yscrollcommand=self.tree_scrollbar.set)

        self.apply_treeview_theme()

        self.split_frame.add(self.tree_frame, stretch="always", minsize=300)

        # --- PANEL PODGLĄDU PDF ---
        self.preview_frame = ctk.CTkFrame(self.split_frame)

        self.lbl_preview_name = ctk.CTkLabel(
            self.preview_frame, text=self.tr("preview_none"),
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            wraplength=280, justify="left"
        )
        self.lbl_preview_name.pack(anchor="w", padx=10, pady=(10, 5))

        # Zwykły tk.Label (nie CTkLabel) dla obrazka podglądu — CTkImage skaluje
        # rozmiar wewnętrznie wg współczynnika DPI, co przy ręcznym dopasowaniu
        # do rozmiaru widgetu dawało podwójne skalowanie i przycinanie na bokach.
        self.lbl_preview_image = tk.Label(self.preview_frame, text="", bd=0, highlightthickness=0)
        self.lbl_preview_image.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Przyciski akcji pod podglądem (pakowane od dołu, więc w odwrotnej
        # kolejności niż wizualna: rząd przycisków na samym dole, nad nim
        # pasek postępu, tuż pod obrazkiem podglądu).
        # grid + uniform gwarantuje RÓWNĄ szerokość obu przycisków niezależnie
        # od długości tekstu (pack z expand=True dzieli tylko nadmiarową
        # przestrzeń, nie całość — stąd nierówne szerokości wcześniej).
        self.btn_row = ctk.CTkFrame(self.preview_frame, fg_color="transparent")
        self.btn_row.pack(side="bottom", fill="x", padx=10, pady=(5, 10))
        self.btn_row.grid_columnconfigure(0, weight=1, uniform="btnrow")
        self.btn_row.grid_columnconfigure(1, weight=1, uniform="btnrow")

        self._adaptive_btn_labels = {}  # widget -> (ikona, pełny_tekst)

        self.btn_open_folder = ctk.CTkButton(self.btn_row, command=self.open_output_folder)
        self.btn_open_folder.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self._set_adaptive_label(self.btn_open_folder, "📂", self.tr("btn_open_folder"))

        self.btn_merge_all = ctk.CTkButton(
            self.btn_row,
            fg_color="green",
            hover_color="darkgreen",
            command=self.merge_all_targets
        )
        self.btn_merge_all.grid(row=0, column=1, sticky="ew", padx=(4, 0))
        self._set_adaptive_label(self.btn_merge_all, "🔄", self.tr("btn_merge"))

        self.btn_row.bind("<Configure>", self._refresh_adaptive_labels)

        self.progress_bar = ctk.CTkProgressBar(self.preview_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(side="bottom", fill="x", padx=10, pady=(0, 5))

        self.split_frame.add(self.preview_frame, stretch="never", minsize=220, width=380)
        self.apply_treeview_theme()  # dociągnij kolory dla lbl_preview_image (nie istniał przy pierwszym wywołaniu)

        self.preview_frame.bind("<Configure>", self.on_preview_resize)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # Wskaźnik linii przeciągania
        self.drop_indicator = tk.Frame(self.tree, height=3, bg="#00a8ff")

        # Rejestracja Drag & Drop z Eksploratora Windows
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<DropPosition>>', self.on_external_drag_over)
        self.tree.dnd_bind('<<DropLeave>>', self.on_external_drag_leave)
        self.tree.dnd_bind('<<Drop>>', self.on_external_drop)

        # Zdarzenia myszy i klawiatury
        self.tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_end)
        self.tree.bind("<Motion>", self.on_mouse_hover)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.tree.bind("<F2>", self.on_f2_key)

        # --- STOPKA PROGRAMU ---
        self.status_frame = ctk.CTkFrame(self, height=28, corner_radius=0)
        self.status_frame.pack(side="bottom", fill="x", padx=0, pady=0)

        self.lbl_status = ctk.CTkLabel(
            self.status_frame,
            text=self.tr("status_ready"),
            anchor="w",
            font=ctk.CTkFont(family="Segoe UI", size=11)
        )
        self.lbl_status.pack(side="left", padx=10, pady=2, fill="x", expand=True)

        self.lbl_website = ctk.CTkLabel(
            self.status_frame,
            text=ABOUT_WEBSITE,
            text_color="#3598db",
            cursor="hand2",
            font=ctk.CTkFont(family="Segoe UI", size=11, underline=True)
        )
        self.lbl_website.pack(side="right", padx=10, pady=2)
        self.lbl_website.bind("<Button-1>", lambda e: webbrowser.open("https://www.draftcon.pl"))

        # Domyślne proporcje okna startowego (850x930 px fizycznych) — 850
        # to tylko dolna granica, żeby rząd przycisków nigdy się nie ucinał;
        # _dpi_geometry() koryguje mnożnik skalowania DPI (patrz wyżej).
        self.update_idletasks()
        start_width = max(self.actions_frame.winfo_reqwidth() + 40, 850)
        self.geometry(self._dpi_geometry(start_width, 930))

        self.refresh_tree()
        self.after(5000, self.periodic_refresh)

    # --- TŁUMACZENIA (PL/EN) ---
    def tr(self, key, **kwargs):
        text = TRANSLATIONS[self.lang].get(key, TRANSLATIONS["pl"].get(key, key))
        return text.format(**kwargs) if kwargs else text

    def toggle_language(self):
        self.lang = "en" if self.lang == "pl" else "pl"
        self._refresh_ui_language()

    # Flaga POKAZUJE aktualnie wybrany język (nie język docelowy przełączenia) —
    # kliknięcie przełącza na drugi język i podmienia flagę na nową bieżącą.
    def _lang_flag_image(self):
        return self._flag_img_pl if self.lang == "pl" else self._flag_img_gb

    def _refresh_ui_language(self):
        self.title(self.tr("app_title", ver=APP_VERSION))
        self.btn_lang.configure(image=self._lang_flag_image())
        self.tt_lang.set_text(self.tr("tt_lang"))

        if self.project_file_path:
            self.lbl_project.configure(text=self.tr("project_label", path=self.project_file_path))
        else:
            self.lbl_project.configure(text=self.tr("project_unsaved"))

        self.tt_load_proj.set_text(self.tr("tt_open_project"))
        self.tt_save_proj.set_text(self.tr("tt_save_project"))
        self.tt_katalog.set_text(self.tr("tt_output_dir"))
        self.tt_new_target.set_text(self.tr("tt_new_target"))
        self.tt_add_pdf.set_text(self.tr("tt_add_pdf"))
        self.tt_refresh.set_text(self.tr("tt_refresh"))
        self.tt_delete.set_text(self.tr("tt_delete"))
        self.tt_undo.set_text(self.tr("tt_undo"))
        self.tt_redo.set_text(self.tr("tt_redo"))
        self.chk_date_suffix.configure(text=self.tr("chk_date_suffix"))

        if self._current_preview_path:
            fname = os.path.basename(self._current_preview_path)
            self.lbl_preview_name.configure(text=self.tr("preview_name", name=fname))
        else:
            self.lbl_preview_name.configure(text=self.tr("preview_none"))

        self._set_adaptive_label(self.btn_open_folder, "📂", self.tr("btn_open_folder"))
        if self._merging:
            self._set_adaptive_label(self.btn_merge_all, "⏳", self.tr("btn_merging"))
        else:
            self._set_adaptive_label(self.btn_merge_all, "🔄", self.tr("btn_merge"))

        self.lbl_status.configure(text=self.tr("status_ready"))
        self.refresh_tree()

    # --- KOREKTA SKALOWANIA DPI DLA self.geometry() ---
    # ctk.CTk.geometry() mnoży podane wymiary przez współczynnik skalowania
    # okna (Windows DPI), więc żeby uzyskać zadany rozmiar FIZYCZNY, trzeba
    # podać wartość podzieloną przez ten współczynnik.
    def _dpi_geometry(self, width, height):
        try:
            scaling = ctk.ScalingTracker.get_window_scaling(self)
        except Exception:
            scaling = 1.0
        return f"{round(width / scaling)}x{round(height / scaling)}"

    # --- AUTOMATYCZNE ODŚWIEŻANIE STANU PLIKÓW W TLE ---
    def periodic_refresh(self):
        if not self._editing:
            self.refresh_tree()
        self.after(5000, self.periodic_refresh)

    # --- UNDO / REDO (stos migawek self.targets) ---
    def _push_undo(self):
        self._undo_stack.append(copy.deepcopy(self.targets))
        if len(self._undo_stack) > self._undo_limit:
            self._undo_stack.pop(0)
        self._redo_stack.clear()
        self._update_undo_redo_buttons()

    def undo(self):
        # Blokada podczas edycji inline (F2/dwuklik na nazwę) — ramka edycji
        # trzyma zamknięcie ze starym t_idx/target; podmiana self.targets pod
        # nią prowadziłaby do zapisania nazwy pod złym indeksem albo IndexError.
        if self._editing or not self._undo_stack:
            return
        self._redo_stack.append(copy.deepcopy(self.targets))
        self.targets = self._undo_stack.pop()
        self._last_moved_iids = []
        self.refresh_tree()
        self._update_undo_redo_buttons()

    def redo(self):
        if self._editing or not self._redo_stack:
            return
        self._undo_stack.append(copy.deepcopy(self.targets))
        self.targets = self._redo_stack.pop()
        self._last_moved_iids = []
        self.refresh_tree()
        self._update_undo_redo_buttons()

    def _update_undo_redo_buttons(self):
        self.btn_undo.configure(state="normal" if self._undo_stack else "disabled")
        self.btn_redo.configure(state="normal" if self._redo_stack else "disabled")

    # --- OTWARCIE FOLDERU WYNIKOWEGO W EXPLORERZE ---
    def open_output_folder(self):
        if not self.output_dir or not os.path.isdir(self.output_dir):
            messagebox.showwarning(self.tr("warn_title"), self.tr("warn_no_output_dir"))
            return
        os.startfile(self.output_dir)

    # --- PRZYCISKI Z ADAPTACYJNYM TEKSTEM (pełny opis albo sama ikona) ---
    def _set_adaptive_label(self, btn, icon, label):
        self._adaptive_btn_labels[btn] = (icon, f"{icon} {label}")
        btn.configure(text=f"{icon} {label}")
        self._refresh_adaptive_labels()

    def _refresh_adaptive_labels(self, event=None):
        if not self._adaptive_btn_labels:
            return
        self.btn_row.update_idletasks()
        row_width = self.btn_row.winfo_width()
        if row_width <= 1:
            return
        # CTkButton + grid(uniform=...) nie zawsze dają dokładnie równą
        # szerokość gdy tekst się zmienia (przycisk potrafi "kurczyć się"
        # do treści) — wymuszamy identyczną szerokość jawnie na obu.
        btn_width = max((row_width - 8) // 2, 40)
        inner_available = btn_width - 20  # margines wewnętrzny przycisku na tekst
        font = tkfont.Font(family="Segoe UI", size=13)
        for btn, (icon, full_text) in self._adaptive_btn_labels.items():
            fits = font.measure(full_text) <= inner_available
            btn.configure(width=btn_width, text=full_text if fits else icon)

    # --- DOSTOSOWANIE KOLORÓW DRZEWA DO TRYBU JASNY/CIEMNY ---
    def apply_treeview_theme(self):
        mode = ctk.get_appearance_mode()  # "Light" lub "Dark"
        if mode == "Dark":
            bg, fg, field_bg, head_bg = "#2b2b2b", "white", "#2b2b2b", "#1f1f1f"
            root_c, target_c = "#3598db", "#ffffff"
            green_c, orange_c, red_c = "#2ecc71", "#ff9f43", "#ff5252"
            sash_bg = "#3a3a3a"
            moved_bg = "#454545"
        else:
            bg, fg, field_bg, head_bg = "#f5f5f5", "#1a1a1a", "#f5f5f5", "#e0e0e0"
            root_c, target_c = "#1a6fb5", "#1a1a1a"
            green_c, orange_c, red_c = "#1e8449", "#c2680a", "#c0392b"
            sash_bg = "#d0d0d0"
            moved_bg = "#d9d9d9"

        if hasattr(self, "split_frame"):
            self.split_frame.configure(bg=sash_bg)
        if hasattr(self, "lbl_preview_image"):
            self.lbl_preview_image.configure(bg=field_bg, fg=fg)

        self.style.configure("Treeview",
                             background=bg,
                             foreground=fg,
                             fieldbackground=field_bg,
                             rowheight=32,
                             font=("Segoe UI", 11))

        self.style.configure("Treeview.Heading",
                             font=("Segoe UI", 11, "bold"),
                             background=head_bg,
                             foreground=fg)

        self.style.map("Treeview", background=[("selected", "#1f538d")])

        self.style.configure("Vertical.TScrollbar",
                             background=head_bg,
                             troughcolor=bg,
                             bordercolor=bg,
                             arrowcolor=fg)

        self.tree.tag_configure("root_branch", foreground=root_c, font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("target_branch", foreground=target_c, font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("file_green", foreground=green_c)
        self.tree.tag_configure("file_orange", foreground=orange_c)
        self.tree.tag_configure("file_red", foreground=red_c)
        self.tree.tag_configure("just_moved", background=moved_bg)

    # --- FORMATOWANIE I TWORZENIE KATALOGU DOCELOWEGO ---
    # Katalog docelowy zostaje zawsze dokładnie taki, jak wpisany ręcznie —
    # data (self.selected_date) NIE jest do niego dopisywana ani synchronizowana
    # przy zmianie daty; dopisywanie daty dotyczy wyłącznie nazw plików wynikowych.
    def process_output_dir_path(self, path):
        if not path:
            return ""

        normalized_path = os.path.normpath(path)

        # Utwórz folder na dysku jeśli nie istnieje
        if not os.path.exists(normalized_path):
            try:
                os.makedirs(normalized_path, exist_ok=True)
            except Exception as e:
                messagebox.showerror(self.tr("create_dir_err_title"), self.tr("create_dir_err_text", err=str(e)))

        return normalized_path

    # --- WYBÓR DATY DOPISYWANEJ JAKO KOŃCÓWKA NAZW ---
    def _strip_date_suffix(self, name):
        return DATE_SUFFIX_RE.sub(r'\2', name)

    def _apply_date_suffix(self, name, date_str):
        base = self._strip_date_suffix(name)
        if base.lower().endswith(".pdf"):
            return f"{base[:-4]}_{date_str}.pdf"
        return f"{base}_{date_str}"

    def open_date_picker(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title(self.tr("date_dialog_title"))
        dialog.transient(self)
        dialog.grab_set()

        cal = Calendar(
            dialog, selectmode="day", date_pattern="yyyy.mm.dd",
            year=self.selected_date.year, month=self.selected_date.month, day=self.selected_date.day
        )
        cal.pack(padx=15, pady=15)

        def confirm():
            y, m, d = (int(v) for v in cal.get_date().split("."))
            dialog.destroy()
            self._on_date_picked(datetime.date(y, m, d))

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(0, 15))
        ctk.CTkButton(btn_frame, text=self.tr("btn_select"), command=confirm).pack(side="right", padx=(8, 0))
        ctk.CTkButton(btn_frame, text=self.tr("btn_cancel"), fg_color="gray", command=dialog.destroy).pack(side="right")

    def _on_date_picked(self, new_date):
        if new_date == self.selected_date:
            return
        self.selected_date = new_date
        new_str = new_date.strftime("%Y.%m.%d")
        self.btn_date_picker.configure(text=f"📅 {new_str}")

        if self.date_suffix_enabled:
            self._push_undo()
            for target in self.targets:
                target["name"] = self._apply_date_suffix(target["name"], new_str)
            self.refresh_tree()

    def _toggle_date_suffix(self):
        self.date_suffix_enabled = self.date_suffix_var.get()
        self.btn_date_picker.configure(state="normal" if self.date_suffix_enabled else "disabled")

        self._push_undo()
        date_str = self.selected_date.strftime("%Y.%m.%d")
        for target in self.targets:
            if self.date_suffix_enabled:
                if not DATE_SUFFIX_RE.search(target["name"]):
                    target["name"] = self._apply_date_suffix(target["name"], date_str)
            else:
                target["name"] = self._strip_date_suffix(target["name"])
        self.refresh_tree()

    # --- OKREŚLANIE AKTYWNEGO ZESTAWU Z ZAZNACZENIA ---
    def get_selected_target_index(self):
        selected = self.tree.selection()
        if not selected:
            return 0
        for item in selected:
            if item.startswith("t_"):
                return int(item.split("_")[1])
            elif item.startswith("f_"):
                return int(item.split("_")[1])
        return 0

    # --- USTAWIENIE WSPÓLNEGO KATALOGU DOCELOWEGO ---
    # Natywny dialog Windows do wyboru folderu odrzuca nieistniejące ścieżki
    # nawet z mustexist=False (ograniczenie Tk), więc własne okno: wpisana
    # ścieżka, jeśli nie istnieje, zostaje po prostu utworzona (jak Explorer).
    def select_output_dir(self, on_confirm=None):
        dialog = ctk.CTkToplevel(self)
        dialog.title(self.tr("outdir_dialog_title"))
        dialog.geometry("600x150")
        dialog.transient(self)
        dialog.grab_set()

        ctk.CTkLabel(
            dialog,
            text=self.tr("outdir_dialog_label")
        ).pack(anchor="w", padx=15, pady=(15, 5))

        entry_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        entry_frame.pack(fill="x", padx=15)

        path_var = tk.StringVar(value=self.output_dir or "")
        entry = ctk.CTkEntry(entry_frame, textvariable=path_var)
        entry.pack(side="left", fill="x", expand=True)

        def browse_existing():
            # Startuj od KATALOGU NADRZĘDNEGO wpisanej ścieżki (np. dla
            # "C:\Projekty\aaa" wejdź do "C:\Projekty") — pokazuje miejsce
            # docelowe z otoczeniem, zamiast wskakiwać w (zwykle pusty)
            # folder, który dopiero powstanie po OK. Jeśli i ten katalog
            # nadrzędny jeszcze nie istnieje, idź wyżej aż do istniejącego.
            start_dir = os.path.dirname(path_var.get().strip())
            while start_dir and not os.path.isdir(start_dir):
                parent = os.path.dirname(start_dir)
                start_dir = parent if parent != start_dir else ""

            kwargs = {"title": self.tr("browse_dialog_title"), "mustexist": True}
            if start_dir:
                kwargs["initialdir"] = start_dir

            picked = filedialog.askdirectory(**kwargs)
            if picked:
                path_var.set(picked)

        ctk.CTkButton(entry_frame, text=self.tr("browse_btn"), width=100, command=browse_existing).pack(side="left", padx=(8, 0))

        def confirm(event=None):
            raw_path = path_var.get().strip()
            if raw_path:
                self.output_dir = self.process_output_dir_path(raw_path)
                self.refresh_tree()
            dialog.destroy()
            if raw_path and on_confirm:
                on_confirm()

        def cancel(event=None):
            dialog.destroy()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=15, side="bottom")
        ctk.CTkButton(btn_frame, text=self.tr("btn_ok"), width=90, command=confirm).pack(side="right", padx=(8, 0))
        ctk.CTkButton(btn_frame, text=self.tr("btn_cancel"), width=90, fg_color="gray", command=cancel).pack(side="right")

        entry.focus_set()
        entry.icursor(tk.END)
        dialog.bind("<Return>", confirm)
        dialog.bind("<Escape>", cancel)

    # --- POKAZYWANIE ŚCIEŻKI W STOPCE ---
    def on_mouse_hover(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            if item == "root":
                self.lbl_status.configure(text=self.tr("status_hover_root", path=self.output_dir or self.tr("root_not_set")))
            elif item.startswith("t_"):
                t_idx = int(item.split("_")[1])
                if t_idx < len(self.targets):
                    target = self.targets[t_idx]
                    full_out = os.path.join(self.output_dir, target['name']) if self.output_dir else target['name']
                    self.lbl_status.configure(text=self.tr("status_hover_target", path=full_out))
            elif item.startswith("f_"):
                parts = item.split("_")
                t_idx, f_idx = int(parts[1]), int(parts[2])
                if t_idx < len(self.targets) and f_idx < len(self.targets[t_idx]["files"]):
                    fpath, _ = self._parse_file_entry(self.targets[t_idx]["files"][f_idx])
                    self.lbl_status.configure(text=self.tr("status_hover_file", path=fpath))
        else:
            self.lbl_status.configure(text=self.tr("status_ready"))

    # --- PODGLĄD PIERWSZEJ STRONY PDF (tylko dla zaznaczenia, cache po mtime) ---
    def on_tree_select(self, event=None):
        selected = self.tree.selection()
        if len(selected) != 1 or not selected[0].startswith("f_"):
            self.show_preview(None)
            return

        parts = selected[0].split("_")
        t_idx, f_idx = int(parts[1]), int(parts[2])
        if t_idx >= len(self.targets) or f_idx >= len(self.targets[t_idx]["files"]):
            self.show_preview(None)
            return

        fpath, _ = self._parse_file_entry(self.targets[t_idx]["files"][f_idx])
        self.show_preview(fpath)

    def show_preview(self, fpath):
        self._current_preview_path = fpath

        if not fpath:
            self.lbl_preview_name.configure(text=self.tr("preview_none"))
            self._current_preview_ctk_img = None
            self.lbl_preview_image.configure(image="", text="")
            return

        fname = os.path.basename(fpath)
        self.lbl_preview_name.configure(text=self.tr("preview_name", name=fname))

        if not os.path.exists(fpath):
            self._current_preview_ctk_img = None
            self.lbl_preview_image.configure(image="", text=self.tr("preview_missing"))
            return

        mtime = os.path.getmtime(fpath)
        cached = self._preview_cache.get(fpath)
        if cached and cached[0] == mtime:
            self._display_preview_image(cached[1])
            return

        try:
            doc = fitz.open(fpath)
            page = doc[0]
            # Renderujemy raz w dość wysokiej rozdzielczości (cache po mtime);
            # dopasowanie do bieżącego rozmiaru panelu robi tanie skalowanie PIL.
            render_w = 700
            zoom = render_w / page.rect.width
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()

            self._preview_cache[fpath] = (mtime, img)
            self._display_preview_image(img)
        except Exception:
            self._current_preview_ctk_img = None
            self.lbl_preview_image.configure(image="", text=self.tr("preview_error"))

    def _display_preview_image(self, pil_img):
        # Rzeczywisty rozmiar widgetu obrazka (nie zgadywany z ramki minus stałe) —
        # nazwa pliku potrafi zawinąć się do 2 linii i zmienić dostępną wysokość.
        self.update_idletasks()
        avail_w = max(self.lbl_preview_image.winfo_width() - 16, 60)
        avail_h = max(self.lbl_preview_image.winfo_height() - 16, 60)

        # Dodatkowy margines bezpieczeństwa (zaokrąglenia) — lepiej lekko za
        # mały obraz niż przycięty na krawędziach.
        scale = min(avail_w / pil_img.width, avail_h / pil_img.height) * 0.96
        disp_w = max(int(pil_img.width * scale), 1)
        disp_h = max(int(pil_img.height * scale), 1)

        resized = pil_img.resize((disp_w, disp_h), Image.LANCZOS)
        photo_img = ImageTk.PhotoImage(resized)
        self._current_preview_ctk_img = photo_img  # trzymaj referencję, inaczej GC skasuje obraz
        self.lbl_preview_image.configure(image=photo_img, text="")

    def on_preview_resize(self, event=None):
        if self._preview_resize_job:
            self.after_cancel(self._preview_resize_job)
        self._preview_resize_job = self.after(120, self._redisplay_current_preview)

    def _redisplay_current_preview(self):
        self._preview_resize_job = None
        fpath = self._current_preview_path
        if not fpath:
            return
        cached = self._preview_cache.get(fpath)
        if cached:
            self._display_preview_image(cached[1])

    # --- SPRAWDZENIE CZY NAZWA PLIKU WYNIKOWEGO JEST JUŻ ZAJĘTA ---
    def _name_exists(self, name, exclude_idx=None):
        return any(
            i != exclude_idx and t["name"].lower() == name.lower()
            for i, t in enumerate(self.targets)
        )

    # --- EDYCJA INLINE ---
    def edit_target_inline(self, item):
        if item == "root":
            self.select_output_dir()
            return

        t_idx = int(item.split("_")[1])
        target = self.targets[t_idx]

        self._editing = True

        # Modalny dialog w stylu select_output_dir() zamiast ciasnej ramki
        # wstawianej bezpośrednio w drzewo — większy, czytelniejszy, bez
        # ograniczeń pozycjonowania względem wiersza.
        dialog = ctk.CTkToplevel(self)
        dialog.title(self.tr("rename_dialog_title"))
        dialog.geometry("320x150")
        dialog.transient(self)
        dialog.grab_set()

        name_var = tk.StringVar(value=target["name"])
        entry = ctk.CTkEntry(dialog, textvariable=name_var, font=ctk.CTkFont(family="Segoe UI", size=13))
        entry.pack(fill="x", padx=15, pady=(20, 0))

        def save_changes(event=None):
            new_name = name_var.get().strip()
            if new_name and not new_name.lower().endswith(".pdf"):
                new_name += ".pdf"
            if new_name and new_name != target["name"]:
                if self._name_exists(new_name, exclude_idx=t_idx):
                    messagebox.showwarning(
                        self.tr("warn_title"),
                        self.tr("name_exists_warn", name=new_name),
                        parent=dialog
                    )
                    return
                self._push_undo()
                self.targets[t_idx]["name"] = new_name
            self._editing = False
            dialog.destroy()
            self.refresh_tree()

        def cancel_changes(event=None):
            self._editing = False
            dialog.destroy()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=15, side="bottom")
        ctk.CTkButton(btn_frame, text=self.tr("btn_ok"), width=90, command=save_changes).pack(side="right", padx=(8, 0))
        ctk.CTkButton(btn_frame, text=self.tr("btn_cancel"), width=90, fg_color="gray", command=cancel_changes).pack(side="right")

        entry.focus_set()
        entry.select_range(0, tk.END)
        entry.icursor(tk.END)

        dialog.bind("<Return>", save_changes)
        dialog.bind("<Escape>", cancel_changes)
        dialog.protocol("WM_DELETE_WINDOW", cancel_changes)

    def on_f2_key(self, event):
        selected = self.tree.selection()
        if selected:
            self.edit_target_inline(selected[0])

    def edit_selected_target(self):
        selected = self.tree.selection()
        if selected:
            self.edit_target_inline(selected[0])

    def on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        if item == "root":
            self.select_output_dir()
        elif item.startswith("t_"):
            self.edit_target_inline(item)
        elif item.startswith("f_"):
            self.open_selected_file(item)

    def open_selected_file(self, item):
        parts = item.split("_")
        t_idx, f_idx = int(parts[1]), int(parts[2])
        if t_idx >= len(self.targets) or f_idx >= len(self.targets[t_idx]["files"]):
            return
        fpath, _ = self._parse_file_entry(self.targets[t_idx]["files"][f_idx])
        if not os.path.exists(fpath):
            messagebox.showerror(self.tr("err_title"), self.tr("err_file_missing", path=fpath))
            return
        os.startfile(fpath)

    # --- ODŚWIEŻANIE DRZEWA ---
    def _parse_file_entry(self, f_entry):
        if isinstance(f_entry, dict):
            return f_entry.get("path", ""), f_entry.get("mtime", None)
        return str(f_entry), None

    def refresh_tree(self):
        expanded_states = {}
        for item in self.tree.get_children():
            expanded_states[item] = self.tree.item(item, "open")
            for child in self.tree.get_children(item):
                expanded_states[child] = self.tree.item(child, "open")

        prev_selection = self.tree.selection()
        prev_focus = self.tree.focus()

        self.tree.delete(*self.tree.get_children())

        out_dir_str = self.output_dir if self.output_dir else self.tr("root_not_set")
        root_label = self.tr("root_label", path=out_dir_str)
        self.tree.insert("", "end", iid="root", text=root_label, open=expanded_states.get("root", True), tags=("root_branch",))

        for t_idx, target in enumerate(self.targets):
            t_iid = f"t_{t_idx}"
            label = f"📁 {target['name']}"

            self.tree.insert("root", "end", iid=t_iid, text=label, open=expanded_states.get(t_iid, True), tags=("target_branch",))

            for f_idx, f_entry in enumerate(target["files"]):
                f_iid = f"f_{t_idx}_{f_idx}"
                fpath, saved_mtime = self._parse_file_entry(f_entry)
                fname = os.path.basename(fpath)

                if not os.path.exists(fpath):
                    tag = "file_red"
                else:
                    curr_mtime = os.path.getmtime(fpath)
                    if saved_mtime is not None and abs(curr_mtime - saved_mtime) > 1.0:
                        tag = "file_orange"
                    else:
                        tag = "file_green"

                file_label = f"    {f_idx + 1:>3}.  📄 {fname}"
                item_tags = (tag, "just_moved") if f_iid in self._last_moved_iids else (tag,)
                self.tree.insert(t_iid, "end", iid=f_iid, text=file_label, tags=item_tags)

        valid_selection = [iid for iid in prev_selection if self.tree.exists(iid)]
        if valid_selection:
            self.tree.selection_set(valid_selection)

        # tree.delete() czyści wewnętrzny "focus" Treeview — a to on jest
        # kotwicą, od której ttk liczy zakres przy Shift-kliknięciu. Bez tego
        # Shift-klik po KAŻDYM refresh_tree() (czyli po prawie każdej akcji
        # i co 5s z periodic_refresh) cichnie degraduje się do zaznaczenia
        # tylko jednego, klikniętego elementu.
        if prev_focus and self.tree.exists(prev_focus):
            self.tree.focus(prev_focus)

    # --- DRAG & DROP MULTI-SELEKCJI WEWNĄTRZ DRZEWA ---
    def on_drag_start(self, event):
        # UWAGA: <ButtonPress-1> odpala się przy KAŻDYM kliknięciu wiersza
        # (także zwykłym zaznaczeniu bez przeciągania) — nie czyścić tu
        # highlightu "ostatnio przesunięty". Ma zniknąć dopiero gdy
        # przeciągnięcie faktycznie się dokończy (on_drag_end/on_external_drop
        # nadpisują _last_moved_iids nowym zestawem).

        # Shift/Ctrl to gest zaznaczania zakresu/przełączania (natywna obsługa
        # ttk.Treeview przy selectmode="extended"), nie przeciąganie do zmiany
        # kolejności — nie uzbrajać dragged_iids, inaczej najmniejszy drgnięcie
        # myszy w trakcie Shift-kliku odpala on_drag_motion/on_drag_end i psuje
        # dopiero co ustawione przez Tk zaznaczenie.
        if event.state & 0x0001 or event.state & 0x0004:  # Shift / Control
            self.dragged_iids = []
            return

        item = self.tree.identify_row(event.y)
        if item and item.startswith("f_"):
            selected = list(self.tree.selection())
            if item in selected:
                self.dragged_iids = [i for i in selected if i.startswith("f_")]
            else:
                self.dragged_iids = [item]
        else:
            self.dragged_iids = []

    # --- AUTO-SCROLL DRZEWA PRZY PRZECIĄGANIU BLISKO GÓRY/DOŁU ---
    # Auto-scroll przesuwa drzewo TIMEREM, bez ruchu myszy — więc cel
    # przeciągania (kreska / _last_drag_target / _last_ext_drop) trzeba
    # przeliczać przy KAŻDYM ticku scrolla, nie tylko przy zdarzeniu ruchu
    # myszy, inaczej zostaje "zamrożony" na starej pozycji sprzed scrolla
    # i drop ląduje w złym miejscu albo nigdzie.
    def _update_autoscroll(self, y, refresh_fn):
        margin = 28
        height = self.tree.winfo_height()
        if y < margin:
            self._set_autoscroll(-1, refresh_fn)
        elif y > height - margin:
            self._set_autoscroll(1, refresh_fn)
        else:
            self._set_autoscroll(None, None)

    def _set_autoscroll(self, direction, refresh_fn):
        self._autoscroll_dir = direction
        self._autoscroll_refresh_fn = refresh_fn
        if direction is not None and self._autoscroll_job is None:
            self._autoscroll_tick()

    def _autoscroll_tick(self):
        if self._autoscroll_dir is None:
            self._autoscroll_job = None
            return
        self.tree.yview_scroll(self._autoscroll_dir, "units")
        if self._autoscroll_refresh_fn:
            self._autoscroll_refresh_fn()
        self._autoscroll_job = self.after(60, self._autoscroll_tick)

    # Wspólna logika wyznaczania celu drop (używana i przy ruchu myszy, i
    # przy każdym ticku auto-scrolla) — działa na "żywej" pozycji kursora
    # pobranej z Tk, nie na współrzędnych ze starego zdarzenia.
    def _compute_drop_target(self, y, exclude_items=()):
        target_item = self.tree.identify_row(y)
        if not target_item or target_item == "root" or target_item in exclude_items:
            return None
        bbox = self.tree.bbox(target_item)
        if not bbox:
            return None
        x, b_y, w, h = bbox

        if target_item.startswith("t_"):
            # Nagłówek gałęzi (zestawu) — zawsze JEDNA linia pod nazwą (bez
            # rozróżniania górnej/dolnej połowy wiersza — mylił przy dwóch
            # pustych gałęziach obok siebie). Znaczenie insert_after tutaj:
            # gałąź PUSTA -> True (jedyna sensowna opcja, dodaj do niej),
            # gałąź NIEPUSTA -> False, czyli "wstaw NA POCZĄTEK listy" (linia
            # ląduje dokładnie pod belką, tuż nad pierwszym plikiem — koniec
            # listy osiąga się przez dolną połowę OSTATNIEGO pliku, jak
            # wcześniej; dawniej nagłówek zawsze znaczył "na koniec", co
            # myliło przy przeciąganiu tuż pod nazwę gałęzi).
            t_idx = int(target_item.split("_")[1])
            has_files = t_idx < len(self.targets) and len(self.targets[t_idx]["files"]) > 0
            return target_item, not has_files, x, b_y, w, h

        insert_after = y >= b_y + (h / 2)
        return target_item, insert_after, x, b_y, w, h

    def _refresh_internal_drag_target(self):
        if not self.dragged_iids:
            return
        y = self.tree.winfo_pointery() - self.tree.winfo_rooty()
        result = self._compute_drop_target(y, exclude_items=self.dragged_iids)
        if result:
            target_item, insert_after, x, b_y, w, h = result
            line_y = b_y + h if insert_after else b_y
            self.drop_indicator.place(x=x, y=line_y, width=w)
            self.drop_indicator.lift()
            self._last_drag_target = (target_item, insert_after)
        else:
            self.drop_indicator.place_forget()
            self._last_drag_target = None

    def on_drag_motion(self, event):
        if not self.dragged_iids:
            return

        self._update_autoscroll(event.y, self._refresh_internal_drag_target)
        self._refresh_internal_drag_target()

    def on_drag_end(self, event):
        self.drop_indicator.place_forget()
        self._set_autoscroll(None, None)

        if not self.dragged_iids:
            return

        # Użyj DOKŁADNIE tej pozycji, którą pokazywała niebieska kreska
        # (a nie przeliczonej na nowo z event.y w momencie puszczenia przycisku)
        drop_info = self._last_drag_target
        self._last_drag_target = None

        if not drop_info:
            self.dragged_iids = []
            return

        drop_target, insert_after = drop_info
        self._push_undo()

        parsed = []
        for iid in self.dragged_iids:
            parts = iid.split("_")
            parsed.append((int(parts[1]), int(parts[2])))

        # Indeks docelowy trzeba policzyć PRZED usunięciem przeciąganych
        # elementów z listy — usunięcie elementów leżących przed celem
        # przesuwa resztę listy, więc "surowy" indeks trzeba skorygować
        # o liczbę usuniętych elementów z tej samej gałęzi, które były
        # przed miejscem wstawienia.
        dest_t_idx = None
        dest_insert_idx = None
        if drop_target.startswith("f_"):
            _, dt_idx, df_idx = drop_target.split("_")
            dest_t_idx = int(dt_idx)
            anchor = int(df_idx) + (1 if insert_after else 0)
            removed_before = sum(1 for t, f in parsed if t == dest_t_idx and f < anchor)
            dest_insert_idx = anchor - removed_before

        parsed.sort(key=lambda x: (x[0], x[1]), reverse=True)

        extracted_entries = []
        for t_idx, f_idx in parsed:
            if t_idx < len(self.targets) and f_idx < len(self.targets[t_idx]["files"]):
                extracted_entries.append(self.targets[t_idx]["files"].pop(f_idx))

        extracted_entries.reverse()

        if drop_target.startswith("t_"):
            dest_t_idx = int(drop_target.split("_")[1])
            if insert_after:
                start_idx = len(self.targets[dest_t_idx]["files"])
                self.targets[dest_t_idx]["files"].extend(extracted_entries)
            else:
                start_idx = 0
                for entry in reversed(extracted_entries):
                    self.targets[dest_t_idx]["files"].insert(0, entry)

        elif drop_target.startswith("f_"):
            start_idx = dest_insert_idx
            for entry in reversed(extracted_entries):
                self.targets[dest_t_idx]["files"].insert(dest_insert_idx, entry)

        self._last_moved_iids = [f"f_{dest_t_idx}_{start_idx + i}" for i in range(len(extracted_entries))]
        self.refresh_tree()
        self.dragged_iids = []

    # --- DRAG & DROP EKSPLORATORA WINDOWS ---
    def _refresh_external_drag_target(self):
        # Celowo NIE używamy event.y_root (tkinterdnd2/natywny Windows OLE
        # drag&drop) razem z winfo_rooty() (Tk) — te dwie wartości potrafią
        # być w różnej przestrzeni współrzędnych (np. przy skalowaniu DPI),
        # co przesuwało trafienie o rząd/dwa. Pytamy Tk o pozycję kursora
        # bezpośrednio, żeby obie wartości pochodziły z tego samego źródła.
        y = self.tree.winfo_pointery() - self.tree.winfo_rooty()
        result = self._compute_drop_target(y)
        if result:
            target_item, insert_after, x, b_y, w, h = result
            line_y = b_y + h if insert_after else b_y
            self.drop_indicator.place(x=x, y=line_y, width=w)
            self.drop_indicator.lift()
            self._last_ext_drop = (target_item, insert_after)
        else:
            self.drop_indicator.place_forget()
            self._last_ext_drop = None
        return y

    def on_external_drag_over(self, event):
        y = self._refresh_external_drag_target()
        self._update_autoscroll(y, self._refresh_external_drag_target)

    def on_external_drag_leave(self, event):
        self.drop_indicator.place_forget()
        self._set_autoscroll(None, None)
        # UWAGA: NIE czyścić self._last_ext_drop tutaj — tkinterdnd2 na
        # Windows potrafi wysłać <<DropLeave>> tuż PRZED właściwym <<Drop>>
        # (kolejność zdarzeń), więc czyszczenie tutaj gubiło zapamiętaną
        # pozycję kreski i drop lądował na końcu listy zamiast w miejscu
        # wskazanym podczas przeciągania.

    def on_external_drop(self, event):
        self.drop_indicator.place_forget()
        self._set_autoscroll(None, None)
        raw_files = self.tk.splitlist(event.data)
        pdf_files = [f.strip('{}') for f in raw_files if f.strip('{}').lower().endswith('.pdf')]

        # Użyj DOKŁADNIE tej pozycji, którą pokazywała niebieska kreska podczas
        # przeciągania (a nie przeliczonej na nowo z pozycji kursora w momencie
        # zdarzenia Drop, która potrafi się nieznacznie różnić)
        drop_info = self._last_ext_drop
        self._last_ext_drop = None

        if not pdf_files:
            return

        t_idx = self.get_selected_target_index()
        insert_f_idx = None

        if drop_info:
            target_item, insert_after = drop_info
            if target_item.startswith("t_"):
                t_idx = int(target_item.split("_")[1])
                if not insert_after:
                    insert_f_idx = 0
            elif target_item.startswith("f_"):
                parts = target_item.split("_")
                t_idx = int(parts[1])
                f_idx = int(parts[2])
                insert_f_idx = f_idx + 1 if insert_after else f_idx

        existing_paths = [self._parse_file_entry(f)[0] for f in self.targets[t_idx]["files"]]

        new_entries = []
        for pdf in pdf_files:
            if pdf not in existing_paths:
                mtime = os.path.getmtime(pdf) if os.path.exists(pdf) else None
                new_entries.append({"path": pdf, "mtime": mtime})

        if new_entries:
            self._push_undo()
            if insert_f_idx is not None:
                start_idx = insert_f_idx
                for entry in reversed(new_entries):
                    self.targets[t_idx]["files"].insert(insert_f_idx, entry)
            else:
                start_idx = len(self.targets[t_idx]["files"])
                self.targets[t_idx]["files"].extend(new_entries)
            self._last_moved_iids = [f"f_{t_idx}_{start_idx + i}" for i in range(len(new_entries))]

        self.refresh_tree()

    # --- USUWANIE ELEMENTÓW ---
    def delete_selected(self):
        selected = self.tree.selection()
        if not selected or "root" in selected:
            return

        targets_to_delete = set()
        files_to_delete = []

        for item in selected:
            if item.startswith("t_"):
                t_idx = int(item.split("_")[1])
                targets_to_delete.add(t_idx)
            elif item.startswith("f_"):
                parts = item.split("_")
                t_idx, f_idx = int(parts[1]), int(parts[2])
                files_to_delete.append((t_idx, f_idx))

        if targets_to_delete:
            if len(self.targets) - len(targets_to_delete) < 1:
                messagebox.showwarning(self.tr("warn_title"), self.tr("warn_min_target"))
                return
            if messagebox.askyesno(self.tr("confirm_title"), self.tr("confirm_delete_targets", n=len(targets_to_delete))):
                self._push_undo()
                for t_idx in sorted(targets_to_delete, reverse=True):
                    del self.targets[t_idx]
                self.refresh_tree()
                return

        if files_to_delete:
            self._push_undo()
            files_to_delete.sort(key=lambda x: (x[0], x[1]), reverse=True)
            for t_idx, f_idx in files_to_delete:
                if t_idx < len(self.targets) and f_idx < len(self.targets[t_idx]["files"]):
                    del self.targets[t_idx]["files"][f_idx]
            self.refresh_tree()

    def show_selected_info(self):
        selected = self.tree.selection()
        if not selected:
            return
        if len(selected) == 1:
            self.display_item_info(selected[0])

    # --- ROZPOZNANIE ROZMIARU STRONY PDF (mm + nazwa formatu, jeśli standardowy) ---
    def _get_page_size_text(self, fpath):
        try:
            doc = fitz.open(fpath)
            rect = doc[0].rect
            doc.close()
        except Exception:
            return self.tr("page_size_error")

        w_mm = rect.width / 72 * 25.4
        h_mm = rect.height / 72 * 25.4
        w_sorted, h_sorted = min(w_mm, h_mm), max(w_mm, h_mm)

        name = None
        for pname, pw, ph in PAPER_SIZES_MM:
            if abs(w_sorted - pw) <= 2 and abs(h_sorted - ph) <= 2:
                name = pname
                break

        orientation = self.tr("orientation_h") if w_mm > h_mm else self.tr("orientation_v")
        extra = ", ".join(x for x in (name, orientation) if x)
        size_str = f"{w_mm:.0f} × {h_mm:.0f} mm"
        return f"{size_str} ({extra})" if extra else size_str

    def display_item_info(self, item):
        if item == "root":
            messagebox.showinfo(self.tr("info_output_dir_title"), self.tr("info_output_dir_body", path=self.output_dir or self.tr("info_not_selected")))
        elif item.startswith("f_"):
            _, t_idx, f_idx = item.split("_")
            t_idx, f_idx = int(t_idx), int(f_idx)
            f_entry = self.targets[t_idx]["files"][f_idx]
            fpath, saved_mtime = self._parse_file_entry(f_entry)
            fname = os.path.basename(fpath)

            if not os.path.exists(fpath):
                status_icon, status_color = "✕", "#e74c3c"
                status_title = self.tr("status_missing_title")
                status_detail = ""
                size_text = "—"
            else:
                size_text = self._get_page_size_text(fpath)
                curr_mtime = os.path.getmtime(fpath)
                curr_dt = datetime.datetime.fromtimestamp(curr_mtime).strftime('%Y-%m-%d %H:%M:%S')
                if saved_mtime is not None and abs(curr_mtime - saved_mtime) > 1.0:
                    saved_dt = datetime.datetime.fromtimestamp(saved_mtime).strftime('%Y-%m-%d %H:%M:%S')
                    status_icon, status_color = "⚠", "#e67e22"
                    status_title = self.tr("status_outdated_title")
                    status_detail = self.tr("detail_saved_disk", saved=saved_dt, curr=curr_dt)
                else:
                    status_icon, status_color = "✓", "#27ae60"
                    status_title = self.tr("status_current_title")
                    status_detail = self.tr("detail_last_modified", curr=curr_dt)

            self._show_file_info_dialog(
                f_idx, fname, size_text, status_icon, status_color, status_title, status_detail, fpath
            )

    # --- OKNO INFORMACJI O PLIKU PDF (kolorowy status z piktogramem) ---
    def _show_file_info_dialog(self, f_idx, fname, size_text, status_icon, status_color, status_title, status_detail, fpath):
        dialog = ctk.CTkToplevel(self)
        dialog.title(self.tr("file_info_dialog_title"))
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)

        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)

        label_font = ctk.CTkFont(family="Segoe UI", size=12, weight="bold")
        value_font = ctk.CTkFont(family="Segoe UI", size=12)

        # Etykieta i wartość w jednej linii po dwukropku — wyjątkiem jest
        # pełna ścieżka (add_wrapped_row), bo bywa za długa na jeden wiersz.
        def add_inline_row(label, value, text_color=None):
            row = ctk.CTkFrame(content, fg_color="transparent")
            row.pack(anchor="w", fill="x", pady=(8, 0))
            ctk.CTkLabel(row, text=label, font=label_font, anchor="w").pack(side="left")
            ctk.CTkLabel(
                row, text=value, font=value_font, text_color=text_color,
                anchor="w", justify="left", wraplength=340
            ).pack(side="left", padx=(6, 0))

        def add_wrapped_row(label, value):
            ctk.CTkLabel(content, text=label, font=label_font, anchor="w").pack(anchor="w", pady=(8, 0))
            ctk.CTkLabel(content, text=value, font=value_font, anchor="w", justify="left", wraplength=420).pack(anchor="w")

        add_inline_row(self.tr("row_index"), str(f_idx + 1))
        add_inline_row(self.tr("row_filename"), fname)
        add_inline_row(self.tr("row_pagesize"), size_text)

        status_row = ctk.CTkFrame(content, fg_color="transparent")
        status_row.pack(anchor="w", fill="x", pady=(8, 0))
        ctk.CTkLabel(status_row, text=self.tr("row_status"), font=label_font, anchor="w").pack(side="left")
        ctk.CTkLabel(
            status_row, text=status_icon,
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=status_color
        ).pack(side="left", padx=(6, 2))
        ctk.CTkLabel(
            status_row, text=status_title, font=label_font,
            text_color=status_color, anchor="w", justify="left", wraplength=320
        ).pack(side="left")
        if status_detail:
            ctk.CTkLabel(
                content, text=status_detail, font=ctk.CTkFont(family="Segoe UI", size=11),
                anchor="w", justify="left", text_color="gray60"
            ).pack(anchor="w", pady=(2, 0))

        add_wrapped_row(self.tr("row_fullpath"), fpath)

        ctk.CTkButton(content, text=self.tr("btn_ok"), width=90, command=dialog.destroy).pack(anchor="e", pady=(16, 0))

        dialog.bind("<Return>", lambda e: dialog.destroy())
        dialog.bind("<Escape>", lambda e: dialog.destroy())

    # --- OKNO "O PROGRAMIE" (freeware, klikalna strona www i e-mail, data wersji) ---
    def show_about_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title(self.tr("about_title"))
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=28, pady=24)

        ctk.CTkLabel(
            content, text=self.tr("about_app_name"),
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            justify="left"
        ).pack(anchor="w")

        ctk.CTkLabel(
            content, text=self.tr("about_version", ver=APP_VERSION),
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color="gray60",
            justify="left"
        ).pack(anchor="w", pady=(0, 12))

        ctk.CTkFrame(content, height=1, fg_color="gray70").pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(
            content, text=self.tr("about_freeware"),
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#2ecc71", justify="left"
        ).pack(anchor="w", pady=(0, 14))

        ctk.CTkLabel(
            content, text=self.tr("about_feedback"),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            justify="left", wraplength=340
        ).pack(anchor="w", pady=(0, 6))

        link_font = ctk.CTkFont(family="Segoe UI", size=12, underline=True)

        lbl_site = ctk.CTkLabel(content, text=ABOUT_WEBSITE, text_color="#3598db", cursor="hand2", font=link_font)
        lbl_site.pack(anchor="w")
        lbl_site.bind("<Button-1>", lambda e: webbrowser.open("https://www.draftcon.pl"))

        lbl_mail = ctk.CTkLabel(content, text=ABOUT_EMAIL, text_color="#3598db", cursor="hand2", font=link_font)
        lbl_mail.pack(anchor="w", pady=(2, 14))
        lbl_mail.bind("<Button-1>", lambda e: webbrowser.open(f"mailto:{ABOUT_EMAIL}"))

        ctk.CTkLabel(
            content, text=self.tr("about_version_date", date=ABOUT_DATE),
            font=ctk.CTkFont(family="Segoe UI", size=11), text_color="gray60",
            justify="left"
        ).pack(anchor="w")

        ctk.CTkButton(content, text=self.tr("btn_ok"), width=90, command=dialog.destroy).pack(anchor="e", pady=(18, 0))

        dialog.bind("<Return>", lambda e: dialog.destroy())
        dialog.bind("<Escape>", lambda e: dialog.destroy())

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        if item not in self.tree.selection():
            self.tree.selection_set(item)

        menu = Menu(self, tearoff=0, font=("Segoe UI", 10))

        if item == "root":
            menu.add_command(label=self.tr("ctx_set_output_dir"), command=self.select_output_dir)
            menu.add_command(label=self.tr("ctx_open_output_folder"), command=self.open_output_folder)
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_new_target"), command=self.add_target)
            menu.add_command(label=self.tr("ctx_refresh"), command=self.refresh_tree)
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_about"), command=self.show_about_dialog)

        elif item.startswith("t_"):
            menu.add_command(label=self.tr("ctx_rename"), command=self.edit_selected_target)
            menu.add_command(label=self.tr("ctx_add_pdf"), command=self.add_files_to_selected)
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_details"), command=self.show_selected_info)
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_new_target"), command=self.add_target)
            menu.add_command(label=self.tr("ctx_refresh"), command=self.refresh_tree)
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_delete_target"), command=self.delete_selected)
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_about"), command=self.show_about_dialog)

        elif item.startswith("f_"):
            menu.add_command(label=self.tr("ctx_open_file"), command=lambda i=item: self.open_selected_file(i))
            menu.add_command(label=self.tr("ctx_reveal"), command=lambda i=item: self.reveal_in_explorer(i))
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_add_pdf"), command=self.add_files_to_selected)
            menu.add_command(label=self.tr("ctx_details"), command=self.show_selected_info)
            selected_files = [i for i in self.tree.selection() if i.startswith("f_")]
            if len(selected_files) >= 2:
                menu.add_separator()
                menu.add_command(label=self.tr("ctx_sort_az"), command=lambda: self.sort_selected(ascending=True))
                menu.add_command(label=self.tr("ctx_sort_za"), command=lambda: self.sort_selected(ascending=False))
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_refresh"), command=self.refresh_tree)
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_delete_file"), command=self.delete_selected)
            menu.add_separator()
            menu.add_command(label=self.tr("ctx_about"), command=self.show_about_dialog)

        menu.post(event.x_root, event.y_root)

    def reveal_in_explorer(self, item):
        parts = item.split("_")
        t_idx, f_idx = int(parts[1]), int(parts[2])
        if t_idx >= len(self.targets) or f_idx >= len(self.targets[t_idx]["files"]):
            return
        fpath, _ = self._parse_file_entry(self.targets[t_idx]["files"][f_idx])
        if not os.path.exists(fpath):
            messagebox.showerror(self.tr("err_title"), self.tr("err_file_missing", path=fpath))
            return
        # Musi być DOKŁADNIE taki format linii poleceń (cudzysłów tylko wokół
        # ścieżki, sklejony z /select, bez spacji) — lista argumentów do Popen
        # owija cały token w cudzysłów, gdy ścieżka ma spację, co explorer.exe
        # błędnie parsuje.
        subprocess.Popen(f'explorer /select,"{os.path.normpath(fpath)}"')

    # --- SORTOWANIE ZAZNACZONYCH PLIKÓW PO NAZWIE ---
    def sort_selected(self, ascending=True):
        selected = [i for i in self.tree.selection() if i.startswith("f_")]
        if len(selected) < 2:
            return

        by_branch = {}
        for iid in selected:
            parts = iid.split("_")
            t_idx, f_idx = int(parts[1]), int(parts[2])
            by_branch.setdefault(t_idx, []).append(f_idx)

        self._push_undo()

        for t_idx, f_idxs in by_branch.items():
            slots = sorted(f_idxs)
            entries = [self.targets[t_idx]["files"][i] for i in slots]
            entries.sort(
                key=lambda e: os.path.basename(self._parse_file_entry(e)[0]).lower(),
                reverse=not ascending
            )
            for slot, entry in zip(slots, entries):
                self.targets[t_idx]["files"][slot] = entry

        self.refresh_tree()

    def add_target(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title(self.tr("new_target_dialog_title"))
        dialog.geometry("480x150")
        dialog.transient(self)
        dialog.grab_set()

        ctk.CTkLabel(
            dialog,
            text=self.tr("new_target_dialog_label")
        ).pack(anchor="w", padx=15, pady=(15, 5))

        name_var = tk.StringVar()
        entry = ctk.CTkEntry(dialog, textvariable=name_var)
        entry.pack(fill="x", padx=15)

        def confirm(event=None):
            name = name_var.get().strip()
            if not name:
                dialog.destroy()
                return
            if not name.lower().endswith(".pdf"):
                name += ".pdf"
            if self._name_exists(name):
                messagebox.showwarning(
                    self.tr("warn_title"),
                    self.tr("name_exists_warn", name=name),
                    parent=dialog
                )
                return
            self._push_undo()
            self.targets.append({"name": name, "files": []})
            self.refresh_tree()
            dialog.destroy()

        def cancel(event=None):
            dialog.destroy()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=15, side="bottom")
        ctk.CTkButton(btn_frame, text=self.tr("btn_ok"), width=90, command=confirm).pack(side="right", padx=(8, 0))
        ctk.CTkButton(btn_frame, text=self.tr("btn_cancel"), width=90, fg_color="gray", command=cancel).pack(side="right")

        entry.focus_set()
        dialog.bind("<Return>", confirm)
        dialog.bind("<Escape>", cancel)

    # --- DODAWANIE PLIKÓW PRZYCISKIEM ---
    def add_files_to_selected(self):
        selected = self.tree.selection()
        t_idx = 0
        insert_f_idx = None

        for item in selected:
            if item.startswith("t_"):
                t_idx = int(item.split("_")[1])
                insert_f_idx = None
                break
            elif item.startswith("f_"):
                parts = item.split("_")
                t_idx = int(parts[1])
                insert_f_idx = int(parts[2]) + 1
                break

        files = filedialog.askopenfilenames(filetypes=[(self.tr("file_filter_pdf"), "*.pdf")])
        if files:
            existing_paths = [self._parse_file_entry(f)[0] for f in self.targets[t_idx]["files"]]
            new_entries = []
            for f in files:
                if f not in existing_paths:
                    mtime = os.path.getmtime(f) if os.path.exists(f) else None
                    new_entries.append({"path": f, "mtime": mtime})

            if new_entries:
                self._push_undo()
                if insert_f_idx is not None:
                    for entry in reversed(new_entries):
                        self.targets[t_idx]["files"].insert(insert_f_idx, entry)
                else:
                    self.targets[t_idx]["files"].extend(new_entries)

            self.refresh_tree()

    def merge_all_targets(self):
        if not self.targets:
            messagebox.showerror(self.tr("err_title"), self.tr("err_no_targets"))
            return

        if not self.output_dir:
            # Zamiast samego błędu — od razu otwórz procedurę wyboru/utworzenia
            # katalogu docelowego; po potwierdzeniu scalanie ruszy automatycznie.
            self.select_output_dir(on_confirm=self.merge_all_targets)
            return

        # Przetworzenie i utworzenie folderu docelowego jeśli nie istnieje
        self.output_dir = self.process_output_dir_path(self.output_dir)

        self._merging = True
        self.btn_merge_all.configure(state="disabled")
        self._set_adaptive_label(self.btn_merge_all, "⏳", self.tr("btn_merging"))
        self.progress_bar.set(0)

        merge_queue = queue.Queue()
        targets_snapshot = [dict(t, files=list(t.get("files", []))) for t in self.targets]
        thread = threading.Thread(
            target=self._merge_worker,
            args=(targets_snapshot, self.output_dir, merge_queue),
            daemon=True
        )
        thread.start()
        self.after(100, self._poll_merge_queue, merge_queue)

    def _merge_worker(self, targets, output_dir, result_queue):
        total = len(targets)
        success_count = 0
        errors = []

        for i, target in enumerate(targets, start=1):
            files = target.get("files", [])
            filename = target.get("name", "")
            if not filename.lower().endswith(".pdf"):
                filename += ".pdf"

            out_path = os.path.join(output_dir, filename)

            if not files:
                errors.append(self.tr("merge_no_files", name=target['name']))
            else:
                try:
                    merger = PdfWriter()
                    for f_entry in files:
                        fpath, _ = self._parse_file_entry(f_entry)
                        if os.path.exists(fpath):
                            merger.append(fpath)
                    merger.write(out_path)
                    merger.close()
                    success_count += 1
                except Exception as e:
                    errors.append(f"- {target['name']}: {str(e)}")
                    if os.path.exists(out_path):
                        try:
                            os.remove(out_path)
                        except OSError:
                            pass

            result_queue.put(("progress", i, total))

        result_queue.put(("done", success_count, total, errors))

    def _poll_merge_queue(self, merge_queue):
        try:
            while True:
                msg = merge_queue.get_nowait()
                if msg[0] == "progress":
                    _, i, total = msg
                    self.progress_bar.set(i / total)
                elif msg[0] == "done":
                    _, success_count, total, errors = msg
                    self.progress_bar.set(0)
                    self._merging = False
                    self.btn_merge_all.configure(state="normal")
                    self._set_adaptive_label(self.btn_merge_all, "🔄", self.tr("btn_merge"))
                    self.refresh_tree()

                    if errors:
                        msg_text = self.tr("merge_result_text", success=success_count, total=total, errors="\n".join(errors))
                        messagebox.showwarning(self.tr("merge_result_title"), msg_text)
                    else:
                        self._show_merge_success_dialog()
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll_merge_queue, merge_queue)

    # --- OKNO SUKCESU PO SCALENIU (z przyciskiem otwierającym folder wynikowy) ---
    def _show_merge_success_dialog(self):
        self._play_success_sound()

        dialog = ctk.CTkToplevel(self)
        dialog.title(self.tr("success_title"))
        dialog.geometry("420x150")
        dialog.transient(self)
        dialog.grab_set()

        ctk.CTkLabel(
            dialog,
            text=self.tr("success_text"),
            font=ctk.CTkFont(family="Segoe UI", size=13),
            wraplength=380,
            justify="left"
        ).pack(anchor="w", padx=15, pady=(20, 5))

        def open_folder_and_close():
            self.open_output_folder()
            dialog.destroy()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=15, side="bottom")
        ctk.CTkButton(btn_frame, text=self.tr("btn_ok"), width=90, command=dialog.destroy).pack(side="right", padx=(8, 0))
        ctk.CTkButton(btn_frame, text=self.tr("open_folder_btn"), width=140, command=open_folder_and_close).pack(side="right")

        dialog.bind("<Return>", lambda e: dialog.destroy())
        dialog.bind("<Escape>", lambda e: dialog.destroy())

    # --- KULTOWY DŹWIĘK WINDOWS "TA-DA" PO UDANYM SCALENIU ---
    def _play_success_sound(self):
        tada_path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Media", "tada.wav")
        try:
            if os.path.exists(tada_path):
                winsound.PlaySound(tada_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
            else:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass

    def save_project(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[(self.tr("file_filter_project"), "*.json")],
            initialdir=os.path.dirname(self.project_file_path) if self.project_file_path else "",
            initialfile=os.path.basename(self.project_file_path) if self.project_file_path else ""
        )
        if path:
            if self.output_dir:
                self.output_dir = self.process_output_dir_path(self.output_dir)

            for target in self.targets:
                updated_files = []
                for f_entry in target["files"]:
                    fpath, _ = self._parse_file_entry(f_entry)
                    mtime = os.path.getmtime(fpath) if os.path.exists(fpath) else None
                    updated_files.append({"path": fpath, "mtime": mtime})
                target["files"] = updated_files

            data = {
                "version": 4.0,
                "output_dir": self.output_dir,
                "targets": self.targets
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)

            self.project_file_path = path
            self.lbl_project.configure(text=self.tr("project_label", path=path))
            self.refresh_tree()
            messagebox.showinfo(self.tr("project_saved_title"), self.tr("project_saved_text"))

    def load_project(self):
        path = filedialog.askopenfilename(
            filetypes=[(self.tr("file_filter_project"), "*.json")],
            initialdir=os.path.dirname(self.project_file_path) if self.project_file_path else ""
        )
        if path and os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.output_dir = data.get("output_dir", "")
            self.targets = data.get("targets", [])
            self.project_file_path = path
            self.lbl_project.configure(text=self.tr("project_label", path=path))
            self._undo_stack.clear()
            self._redo_stack.clear()
            self._update_undo_redo_buttons()
            self.refresh_tree()

if __name__ == "__main__":
    app = PDFMultiMergerApp()
    app.mainloop()
