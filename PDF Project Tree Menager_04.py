"""PDF Project Tree Manager — wersja na PySide6 + PyQt-Fluent-Widgets.

Port interfejsu z tkinter/customtkinter na Qt. Logika (model danych, scalanie,
format projektu .json, sufiksy dat) przeniesiona 1:1 — zmienia się wyłącznie
warstwa GUI.

Zyski wynikające z samego przejścia na Qt:
  * zoom/przesuwanie podglądu robi QGraphicsView (koniec ręcznego kadrowania),
  * wskaźnik wstawiania i autoscroll przy przeciąganiu daje QTreeWidget,
  * ikony są wektorowe (FluentIcon), więc dają się kolorować — inaczej niż emoji,
  * skalowanie DPI obsługuje Qt (nie trzeba korygować geometrii okna),
  * znika Pillow (render PDF idzie prosto do QImage) i babel wraz z tkcalendar.
"""

import copy
import datetime
import json
import os
import re
import subprocess
import sys
import winreg
import winsound

import pymupdf as fitz
from pypdf import PdfWriter

from PySide6.QtCore import (QDate, QEvent, QItemSelectionModel, QPoint, QRect, QRectF, QSize,
                            Qt, QThread, QTimer, Signal)
from PySide6.QtGui import (QAction, QColor, QDrag, QFontMetrics, QGuiApplication,
                           QIcon, QImage, QPainter, QPen, QPixmap)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QFileDialog,
                               QSplitterHandle, QStyle,
                               QGraphicsPixmapItem, QGraphicsScene, QGraphicsView,
                               QHBoxLayout, QMainWindow, QMenu, QSizePolicy,
                               QSplitter, QStatusBar, QTreeWidgetItem, QVBoxLayout,
                               QWidget)

from qfluentwidgets import (BodyLabel, CalendarPicker, CaptionLabel, CheckBox, ComboBox,
                            FluentIcon, InfoBadge, InfoBar, InfoBarPosition, LineEdit,
                            MessageBoxBase, PrimaryPushButton, ProgressBar, PushButton,
                            SpinBox,
                            StrongBodyLabel, SubtitleLabel, SwitchButton, Theme,
                            ToggleToolButton, ToolButton,
                            ToolTipFilter, ToolTipPosition, TransparentToolButton,
                            TreeItemDelegate, TreeWidget, isDarkTheme, setTheme,
                            setThemeColor, themeColor)

DATE_SUFFIX_RE = re.compile(r'_(\d{4}\.\d{2}\.\d{2})(\.pdf)$', re.IGNORECASE)
APP_VERSION = "0.4"
ABOUT_WEBSITE = "WWW.DRAFTCON.PL"
ABOUT_EMAIL = "biuro@draftcon.pl"
ABOUT_DATE = "09.2026"

PAPER_SIZES_MM = [
    ("A0", 841, 1189), ("A1", 594, 841), ("A2", 420, 594),
    ("A3", 297, 420), ("A4", 210, 297), ("A5", 148, 210),
    ("Letter", 216, 279), ("Legal", 216, 356),
]

PREVIEW_RENDER_WIDTH = 1500
PREVIEW_CACHE_LIMIT = 12
PREVIEW_ZOOM_MAX = 8.0
PREVIEW_ZOOM_MIN = 1.0

PROJECT_HISTORY_FILENAME = "project_history.json"
PROJECT_HISTORY_LIMIT = 15

# Zapamiętany rozmiar okna i pozycja suwaka — plik leży obok programu,
# tak samo jak historia projektów (patrz get_app_dir).
WINDOW_STATE_FILENAME = "window_state.json"
DEFAULT_WINDOW_SIZE = (610, 510)
DEFAULT_SPLIT = (330, 265)
MIN_WINDOW_SIZE = (520, 400)
# Ponizej tej szerokosci panelu dolne przyciski trace napisy i zostaja same
# ikony — inaczej ich napisy (135 + 127 px) tworza podloge, o ktora zatrzymuje
# sie prawa krawedz okna i podgladu nie da sie juz zwezic.
PREVIEW_COMPACT_WIDTH = 300
# Twarde minimum kalendarza — nizsze niz jego sizeHint, wiec przy waskim oknie
# sam sie zweza zamiast blokowac zwezanie okna. Checkbox daty jest bez napisu
# (sam kwadrat z dymkiem) — wczesniej jego napis razem z kalendarzem tworzyl
# podloge 581 px i okno odbijalo przy zwezaniu.
DATE_PICKER_MIN_WIDTH = 104
# Przelacznik daty bez napisow ma sizeHint 56 px (miejsce na On/Off), a sam
# suwak potrzebuje 44 — nadmiar podnosilby podloge okna ponad MIN_WINDOW_SIZE.
DATE_SWITCH_MIN_WIDTH = 44
# Numeracja stron w scalonym PDF (opcja, domyślnie wyłączona).
PAGE_NUMBER_POSITIONS = ("left", "center", "right")
PAGE_NUMBER_DEFAULTS = {"position": "center", "offset_x_mm": 0, "offset_y_mm": 7,
                        "prefix": "STRONA", "font_size": 9}
PAGE_NUMBER_PREFIX_MAX = 40   # znaków; dłuższy tekst i tak nie zmieści się na stronie
PAGE_NUMBER_OFFSET_MAX = 100  # mm; więcej i numer wyjdzie poza kartkę
PAGE_NUMBER_FONT_MIN = 5
PAGE_NUMBER_FONT_MAX = 48
# Odsunięcia podajemy w MILIMETRACH, bo tak opisuje się dokumentację budowlaną.
# Wnętrze PDF-a liczy w punktach (1 pt = 1/72 cala), więc przeliczamy przy
# rysowaniu. Rozmiar czcionki zostaje w punktach — to jednostka typograficzna.
MM_TO_POINTS = 72.0 / 25.4
# Autoprzewijanie przy przeciąganiu: strefa przy krawędzi listy, tempo
# odświeżania i zakres kroku (px na tik — od wejścia w strefę do jej końca).
DRAG_SCROLL_MARGIN = 36
DRAG_SCROLL_INTERVAL = 30
DRAG_SCROLL_MIN_STEP = 4
DRAG_SCROLL_MAX_STEP = 26
# Uchwyt podziału: szerokość strefy chwytania i pionowy margines kreski.
SPLIT_HANDLE_WIDTH = 9
SPLIT_LINE_MARGIN = 6

# Role danych w drzewie: przechowujemy pozycję w modelu (self.targets) wprost
# na elemencie, zamiast kodować ją w tekstowym iid jak w wersji tkinter.
ROLE_KIND = Qt.UserRole          # "root" | "target" | "file"
ROLE_TARGET = Qt.UserRole + 1    # indeks zestawu
ROLE_FILE = Qt.UserRole + 2      # indeks pliku w zestawie
ROLE_NAME = Qt.UserRole + 3      # sama nazwa pliku (tekst w drzewie ma prefiks z numerem)


# --- KATALOG APLIKACJI (do pliku historii projektów) ---
# Każdy tryb "onefile" rozpakowuje się do katalogu tymczasowego, więc __file__
# wskazywałby na kasowany temp. PyInstaller ustawia sys.frozen (a sys.executable
# pokazuje prawdziwy .exe), Nuitka definiuje __compiled__ i NUITKA_ONEFILE_BINARY.
def get_resource_dir():
    """Katalog z zasobami DOLACZONYMI do programu (np. app.ico).

    W trybie onefile PyInstaller rozpakowuje je do katalogu tymczasowego
    (sys._MEIPASS) — to NIE jest miejsce na ustawienia, te maja lezec obok
    pliku wykonywalnego, czyli w get_app_dir().
    """
    return getattr(sys, "_MEIPASS", None) or get_app_dir()


def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    if "__compiled__" in globals():
        exe_path = os.environ.get("NUITKA_ONEFILE_BINARY") or sys.executable
        return os.path.dirname(os.path.abspath(exe_path))
    return os.path.dirname(os.path.abspath(__file__))


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
        "tt_on_top": "Zawsze na wierzchu",
        "tt_on_top_off": "Zawsze na wierzchu: wyłączone",
        "tt_prev_page": "Poprzednia strona",
        "tt_next_page": "Następna strona",
        "preview_hint": "Kółko = strony · Ctrl+kółko = powiększenie\nPrzeciąganie = przesuwanie · Dwuklik = cała strona",
        "ctx_help": "Pomoc...",
        "help_title": "Skrócona instrukcja",
        "help_add_t": "Dodawanie",
        "help_add": "Przeciągnij pliki PDF z Eksploratora na zestaw albo użyj ➕.",
        "help_order_t": "Kolejność",
        "help_order": "Przeciągaj pliki w drzewie — linia pokazuje miejsce wstawienia.",
        "menu_invert_wheel": "Odwróć kółko myszy",
        "menu_page_numbers": "Numeracja stron...",
        "pn_title": "Numeracja stron scalonego PDF",
        "pn_enable": "Numeruj strony",
        "pn_first": "Numer na pierwszej stronie",
        "pn_position": "Położenie",
        "pn_left": "Lewa",
        "pn_center": "Środek",
        "pn_right": "Prawa",
        "pn_prefix": "Przedrostek",
        "pn_offset_x": "Odsunięcie w poziomie",
        "pn_offset_y": "Odsunięcie od dołu",
        "pn_font": "Rozmiar czcionki",
        "pn_unit_mm": " mm",
        "pn_unit_pt": " pkt",
        "pn_hint_edge": "Odsunięcie w poziomie liczone od krawędzi strony.",
        "pn_hint_center": "Przy położeniu na środku odsunięcie w poziomie przesuwa numer względem osi (0 = dokładnie na środku).",
        "pn_format": "Format: {example}",
        "help_preview_t": "Podgląd",
        "help_preview": "Kółko = powiększenie, przeciąganie = przesuwanie, dwuklik = cała strona, Shift+kółko = strony.",
        "help_dots_t": "Kropki przy plikach",
        "help_dots": "Zielona = aktualny, pomarańczowa = zmieniony po zapisie, czerwona = brak pliku. Wypełniona = ostatnia operacja.",
        "tt_unsaved": "Projekt ma niezapisane zmiany",
        "chk_date_suffix": "Dopisuj datę",
        "preview_name": "Podgląd: {name}",
        "preview_missing": "Plik nie istnieje na dysku",
        "preview_error": "Nie można wygenerować podglądu (uszkodzony PDF?)",
        "btn_open_folder": "Otwórz folder",
        "btn_merge": "Scal zestawy",
        "btn_merging": "Scalanie...",
        "status_ready": "Gotowy",
        "status_hover_root": "Wspólny katalog docelowy: {path}",
        "status_hover_target": "Docelowy plik PDF: {path}",
        "status_hover_file": "{path}",
        "warn_title": "Uwaga",
        "warn_no_output_dir": "Katalog docelowy nie jest jeszcze ustawiony lub nie istnieje.",
        "err_title": "Błąd",
        "err_file_missing": "Plik nie istnieje na dysku:\n{path}",
        "btn_cancel": "Anuluj",
        "outdir_dialog_title": "Katalog docelowy",
        "outdir_dialog_label": "Ścieżka katalogu docelowego (jeśli nie istnieje, zostanie utworzona):",
        "browse_btn": "Przeglądaj...",
        "browse_dialog_title": "Przeglądaj istniejące katalogi",
        "btn_ok": "OK",
        "name_exists_warn": "Plik wynikowy o nazwie „{name}” już istnieje w projekcie.",
        "rename_dialog_title": "Nazwa pliku wynikowego",
        # Bez prefiksu "KATALOG DOCELOWY:" — ikona folderu i pozycja korzenia
        # mówią to samo, a prefiks tylko zabierał miejsce długiej ścieżce.
        "root_label": "{path}",
        "root_not_set": "NIE USTAWIONO KATALOGU DOCELOWEGO",
        "page_size_error": "nie udało się odczytać",
        "orientation_h": "poziomo",
        "orientation_v": "pionowo",
        "status_missing_title": "PLIK NIE ISTNIEJE NA DYSKU",
        "status_outdated_title": "NIEAKTUALNY (zmodyfikowany po zapisaniu)",
        "status_current_title": "AKTUALNY (bez zmian)",
        "detail_saved_disk": "Data w projekcie: {saved}\nData na dysku: {curr}",
        "detail_last_modified": "Data ostatniej modyfikacji: {curr}",
        "file_info_dialog_title": "Informacje o pliku PDF",
        "row_index": "Numer w zestawieniu:",
        "row_filename": "Nazwa pliku:",
        "row_pagesize": "Rozmiar strony:",
        "row_pages": "Liczba stron:",
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
        "about_app_name": "PDF Project Tree Manager",
        "about_version": "wersja {ver}",
        "about_license": "Wolne oprogramowanie na licencji GNU AGPL v3",
        "about_source": "Kod źródłowy dostępny na GitHub",
        "about_components": "Wykorzystuje: PyMuPDF (AGPL-3.0), PySide6 (LGPL-3.0), PySide6-Fluent-Widgets (GPL-3.0), pypdf (BSD-3).",
        "about_feedback": "Aby podzielić się opinią lub otrzymać nowszą wersję programu, napisz:",
        "about_version_date": "Wersja z: {date}",
        "confirm_title": "Potwierdzenie",
        "confirm_delete_targets": "Czy na pewno chcesz usunąć zaznaczone zestawy ({n}) wraz z plikami?",
        "err_no_targets": "Brak zestawów do scalenia.",
        "warn_min_target": "Projekt musi posiadać co najmniej jeden plik wynikowy.",
        "new_target_dialog_title": "Nowy plik wynikowy",
        "new_target_dialog_label": "Wprowadź nazwę pliku wyjściowego (np. Operat_Projektowy.pdf):",
        "merge_result_title": "Wynik scalania",
        "merge_result_text": "Zaktualizowano zestawów: {success} / {total}\n\nProblemy:\n{errors}",
        "merge_no_files": "- {name}: Brak plików składowych.",
        "success_title": "Sukces",
        "success_text": "Pomyślnie zaktualizowano wszystkie pliki wyjściowe!",
        "open_folder_btn": "Otwórz folder",
        "create_dir_err_title": "Błąd tworzenia katalogu",
        "create_dir_err_text": "Nie udało się utworzyć katalogu:\n{err}",
        "file_filter_pdf": "Pliki PDF (*.pdf)",
        "file_filter_project": "Projekt PDF Merger (*.json)",
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
        "tt_on_top": "Always on top",
        "tt_on_top_off": "Always on top: off",
        "tt_prev_page": "Previous page",
        "tt_next_page": "Next page",
        "preview_hint": "Wheel = pages · Ctrl+wheel = zoom\nDrag = pan · Double-click = fit whole page",
        "ctx_help": "Help...",
        "help_title": "Quick guide",
        "help_add_t": "Adding",
        "help_add": "Drag PDF files from Explorer onto a set, or use ➕.",
        "help_order_t": "Order",
        "help_order": "Drag files within the tree — the line shows the insertion point.",
        "menu_invert_wheel": "Invert mouse wheel",
        "menu_page_numbers": "Page numbering...",
        "pn_title": "Page numbering of the merged PDF",
        "pn_enable": "Number pages",
        "pn_first": "Number on the first page",
        "pn_position": "Position",
        "pn_left": "Left",
        "pn_center": "Center",
        "pn_right": "Right",
        "pn_prefix": "Prefix",
        "pn_offset_x": "Horizontal offset",
        "pn_offset_y": "Offset from the bottom",
        "pn_font": "Font size",
        "pn_unit_mm": " mm",
        "pn_unit_pt": " pt",
        "pn_hint_edge": "Horizontal offset is measured from the page edge.",
        "pn_hint_center": "With the center position, the horizontal offset shifts the number off the axis (0 = exactly centered).",
        "pn_format": "Format: {example}",
        "help_preview_t": "Preview",
        "help_preview": "Wheel = zoom, drag = pan, double-click = fit whole page, Shift+wheel = pages.",
        "help_dots_t": "Dots next to files",
        "help_dots": "Green = up to date, orange = changed after saving, red = missing. Filled = last operation.",
        "tt_unsaved": "The project has unsaved changes",
        "chk_date_suffix": "Append date",
        "preview_name": "Preview: {name}",
        "preview_missing": "File does not exist on disk",
        "preview_error": "Unable to generate preview (corrupted PDF?)",
        "btn_open_folder": "Open folder",
        "btn_merge": "Merge sets",
        "btn_merging": "Merging...",
        "status_ready": "Ready",
        "status_hover_root": "Shared output directory: {path}",
        "status_hover_target": "Target PDF file: {path}",
        "status_hover_file": "{path}",
        "warn_title": "Warning",
        "warn_no_output_dir": "The output directory is not set yet or does not exist.",
        "err_title": "Error",
        "err_file_missing": "File does not exist on disk:\n{path}",
        "btn_cancel": "Cancel",
        "outdir_dialog_title": "Output directory",
        "outdir_dialog_label": "Output directory path (will be created if it doesn't exist):",
        "browse_btn": "Browse...",
        "browse_dialog_title": "Browse existing directories",
        "btn_ok": "OK",
        "name_exists_warn": "An output file named \u201c{name}\u201d already exists in the project.",
        "rename_dialog_title": "Output file name",
        "root_label": "{path}",
        "root_not_set": "OUTPUT DIRECTORY NOT SET",
        "page_size_error": "could not be read",
        "orientation_h": "landscape",
        "orientation_v": "portrait",
        "status_missing_title": "FILE DOES NOT EXIST ON DISK",
        "status_outdated_title": "OUTDATED (modified after saving)",
        "status_current_title": "UP TO DATE (unchanged)",
        "detail_saved_disk": "Date in project: {saved}\nDate on disk: {curr}",
        "detail_last_modified": "Last modified: {curr}",
        "file_info_dialog_title": "PDF file information",
        "row_index": "Number in the set:",
        "row_filename": "File name:",
        "row_pagesize": "Page size:",
        "row_pages": "Page count:",
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
        "about_app_name": "PDF Project Tree Manager",
        "about_version": "version {ver}",
        "about_license": "Free software under the GNU AGPL v3 license",
        "about_source": "Source code available on GitHub",
        "about_components": "Uses: PyMuPDF (AGPL-3.0), PySide6 (LGPL-3.0), PySide6-Fluent-Widgets (GPL-3.0), pypdf (BSD-3).",
        "about_feedback": "To share feedback or get a newer version of the program, write to:",
        "about_version_date": "Version date: {date}",
        "confirm_title": "Confirmation",
        "confirm_delete_targets": "Are you sure you want to delete the selected sets ({n}) along with their files?",
        "err_no_targets": "No sets to merge.",
        "warn_min_target": "The project must have at least one output file.",
        "new_target_dialog_title": "New output file",
        "new_target_dialog_label": "Enter the output file name (e.g. Project_Report.pdf):",
        "merge_result_title": "Merge result",
        "merge_result_text": "Sets updated: {success} / {total}\n\nIssues:\n{errors}",
        "merge_no_files": "- {name}: No component files.",
        "success_title": "Success",
        "success_text": "All output files were updated successfully!",
        "open_folder_btn": "Open folder",
        "create_dir_err_title": "Directory creation error",
        "create_dir_err_text": "Failed to create the directory:\n{err}",
        "file_filter_pdf": "PDF files (*.pdf)",
        "file_filter_project": "PDF Merger project (*.json)",
    },
}


# --- FLAGI PL/GB RYSOWANE QPAINTEREM ---
# Jak w wersji tkinter: Windows nie rysuje emoji flag jako obrazków, więc
# rysujemy je sami. Tu jednak bez Pillow — prosto na QPixmap.
def _flag_pixmap(lang, w=32, h=20):
    pix = QPixmap(w, h)
    painter = QPainter(pix)
    painter.setRenderHint(QPainter.Antialiasing)

    if lang == "pl":
        painter.fillRect(0, 0, w, h // 2, QColor("white"))
        painter.fillRect(0, h // 2, w, h - h // 2, QColor("#dc143c"))
    else:
        painter.fillRect(0, 0, w, h, QColor("#00247d"))
        painter.setPen(QPen(QColor("white"), max(int(h * 0.22), 2)))
        painter.drawLine(0, 0, w, h)
        painter.drawLine(0, h, w, 0)
        painter.setPen(QPen(QColor("#cf142b"), max(int(h * 0.10), 1)))
        painter.drawLine(0, 0, w, h)
        painter.drawLine(0, h, w, 0)
        cross = max(int(h * 0.34), 3)
        painter.fillRect(w // 2 - cross // 2, 0, cross, h, QColor("white"))
        painter.fillRect(0, h // 2 - cross // 2, w, cross, QColor("white"))
        cross2 = max(int(h * 0.18), 2)
        painter.fillRect(w // 2 - cross2 // 2, 0, cross2, h, QColor("#cf142b"))
        painter.fillRect(0, h // 2 - cross2 // 2, w, cross2, QColor("#cf142b"))

    painter.setPen(QPen(QColor("#999999"), 1))
    painter.drawRect(0, 0, w - 1, h - 1)
    painter.end()
    return pix


# --- KROPKI STANU PLIKU ---
# Zamiast kolorować tekst nazw plików, stan niesie kropka przed nazwą:
# kolor = aktualność pliku, wypełnienie = element ostatniej operacji.
# Tekst zostaje w domyślnym kolorze, więc nazwy czytają się lepiej.
_DOT_CACHE = {}


def page_number_text(prefix, page, total):
    """Napis numeru strony: przedrostek (jeśli jest) + strona/liczba stron."""
    prefix = (prefix or "").strip()
    number = f"{page}/{total}"
    return f"{prefix} {number}" if prefix else number


def _dot_icon(color_name, filled, size=16):
    key = (color_name, filled, size)
    if key in _DOT_CACHE:
        return _DOT_CACHE[key]

    app = QApplication.instance()
    dpr = app.primaryScreen().devicePixelRatio() if app else 1.0
    pixmap = QPixmap(int(size * dpr), int(size * dpr))
    pixmap.setDevicePixelRatio(dpr)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    color = QColor(color_name)
    margin = 4.0 if filled else 4.5
    rect = QRectF(margin, margin, size - 2 * margin, size - 2 * margin)
    if filled:
        painter.setPen(Qt.NoPen)
        painter.setBrush(color)
    else:
        painter.setPen(QPen(color, 1.8))  # cieńszy pierścień gasił kolor
        painter.setBrush(Qt.NoBrush)
    painter.drawEllipse(rect)
    painter.end()

    icon = QIcon(pixmap)
    _DOT_CACHE[key] = icon
    return icon


class HistoryComboBox(ComboBox):
    """Lista projektów, która weryfikuje wpisy tuż przed rozwinięciem."""

    aboutToShow = Signal()

    def showPopup(self):
        self.aboutToShow.emit()
        super().showPopup()


class SplitHandle(QSplitterHandle):
    """Uchwyt podziału z widoczną, cienką kreską.

    Styl Fluent rysuje uchwyt na przezroczysto, więc miejsce chwytania było
    niewidoczne — trzeba go było szukać po zmianie kursora. Kreska pokazuje,
    gdzie ono jest; pod kursorem rozjaśnia się kolorem motywu.
    """

    def __init__(self, orientation, parent):
        super().__init__(orientation, parent)
        self.setCursor(Qt.SplitHCursor if orientation == Qt.Horizontal
                       else Qt.SplitVCursor)

    def enterEvent(self, event):
        super().enterEvent(event)
        self.update()

    def leaveEvent(self, event):
        super().leaveEvent(event)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        hovered = self.underMouse()
        if hovered:
            color = themeColor()
        else:
            color = QColor("#4a4a4a") if isDarkTheme() else QColor("#c8c8c8")
        thickness = 2 if hovered else 1
        if self.orientation() == Qt.Horizontal:
            x = (self.width() - thickness) // 2
            painter.fillRect(x, SPLIT_LINE_MARGIN, thickness,
                             max(self.height() - 2 * SPLIT_LINE_MARGIN, 0), color)
        else:
            y = (self.height() - thickness) // 2
            painter.fillRect(SPLIT_LINE_MARGIN, y,
                             max(self.width() - 2 * SPLIT_LINE_MARGIN, 0), thickness, color)
        painter.end()


class SplitView(QSplitter):
    """Podział okna, który używa uchwytu z widoczną kreską."""

    def createHandle(self):
        return SplitHandle(self.orientation(), self)


class PreviewPanel(QWidget):
    """Panel podgladu — zglasza swoja szerokosc, zeby dolne przyciski
    mogly przelaczac sie na same ikony i nie blokowac zwezania okna."""

    resized = Signal(int)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.resized.emit(self.width())


class ElidedLabel(StrongBodyLabel):
    """Etykieta jednowierszowa, która NIGDY nie rozpycha układu.

    Zwykła etykieta zgłasza szerokość równą długości tekstu, więc długa nazwa
    pliku dusiła drzewo (305 -> 79 px) i rozciągała okno główne (610 -> 782 px)
    — zmierzone. Tutaj szerokość podpowiedzi to 0, a tekst jest skracany
    wielokropkiem do aktualnej szerokości; pełna treść zostaje w dymku.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._full = ""
        self.setWordWrap(False)
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)

    def setText(self, text):
        self._full = text or ""
        self.setToolTip(self._full)
        self._apply_elide()

    def _apply_elide(self):
        width = max(self.width(), 0)
        if width <= 0:
            super().setText(self._full)
            return
        super().setText(QFontMetrics(self.font()).elidedText(
            self._full, Qt.ElideMiddle, width))

    def sizeHint(self):
        # Szerokosc 0 — o szerokosci decyduje wylacznie panel, nie tekst.
        return QSize(0, super().sizeHint().height())

    def minimumSizeHint(self):
        return QSize(0, super().sizeHint().height())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_elide()


class PreviewView(QGraphicsView):
    """Podgląd strony PDF z powiększaniem i przesuwaniem.

    Cała mechanika, która w wersji tkinter wymagała ręcznego kadrowania obrazu
    (crop + resize + clamp offsetu), jest tutaj wbudowana w QGraphicsView:
    przesuwanie to ScrollHandDrag, a powiększanie to scale() wokół kursora.
    """

    wheelPaged = Signal(int)  # Shift+kółko przy pliku wielostronicowym

    def __init__(self, parent=None):
        super().__init__(parent)
        self._scene = QGraphicsScene(self)
        self.setScene(self._scene)
        self._item = QGraphicsPixmapItem()
        self._scene.addItem(self._item)

        self.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setAlignment(Qt.AlignCenter)
        # Ramka i białe tło są widoczne ZAWSZE — także gdy nic nie wybrano,
        # żeby panel podglądu był czytelnym, pustym obszarem, a nie dziurą.
        self.setFrameShape(QGraphicsView.StyledPanel)
        self.setBackgroundBrush(QColor("#2b2b2b") if isDarkTheme() else QColor("#ffffff"))
        self._placeholder = ""
        self._placeholder_error = False
        # Rozmiar widoku ma wynikać WYŁĄCZNIE z okna, nigdy z zawartości.
        # QGraphicsView domyślnie wylicza sizeHint ze sceny, więc po wczytaniu
        # strony zgłaszał układowi "chcę być szeroki jak obrazek" i rozpychał
        # panel. Kierunek ma być odwrotny: obrazek dopasowuje się do panelu.
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self._zoom = 1.0  # 1.0 = CAŁA strona widoczna w panelu
        # Kierunek kółka: domyślnie obrót OD SIEBIE powiększa. Przełączalny
        # we wspólnym menu pod prawym przyciskiem, bo to kwestia
        # przyzwyczajenia — różne czytniki PDF robią to różnie.
        self._wheel_inverted = False
        self._has_image = False
        self._multipage = False

    def wheel_inverted(self):
        return self._wheel_inverted

    def set_wheel_inverted(self, inverted):
        self._wheel_inverted = bool(inverted)

    def set_placeholder(self, text, error=False):
        """Tekst rysowany na środku pustego podglądu (podpowiedź albo błąd)."""
        self._placeholder = text or ""
        self._placeholder_error = error
        self.viewport().update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._has_image or not self._placeholder:
            return
        painter = QPainter(self.viewport())
        if self._placeholder_error:
            painter.setPen(QColor("#ff5252") if isDarkTheme() else QColor("#c0392b"))
        else:
            painter.setPen(QColor("#9a9a9a") if isDarkTheme() else QColor("#8a8a8a"))
        # Margines i zawijanie — w wąskim panelu tekst inaczej dotyka krawędzi
        # ramki i jest przycinany po bokach.
        rect = self.viewport().rect().adjusted(14, 14, -14, -14)
        painter.drawText(rect, Qt.AlignCenter | Qt.TextWordWrap, self._placeholder)
        painter.end()

    def sizeHint(self):
        # Stała, skromna podpowiedź rozmiaru — niezależna od wczytanej strony.
        return QSize(240, 240)

    def minimumSizeHint(self):
        return QSize(120, 120)

    def set_pixmap(self, pixmap, keep_view=False):
        self._item.setPixmap(pixmap)
        self._scene.setSceneRect(self._item.boundingRect())
        self._has_image = not pixmap.isNull()
        if not keep_view:
            self.reset_view()

    def clear_image(self):
        self._item.setPixmap(QPixmap())
        self._has_image = False
        self._zoom = 1.0
        self.resetTransform()

    def set_multipage(self, multipage):
        self._multipage = multipage

    def reset_view(self):
        self.fit_page()

    def fit_page(self):
        """Skala bazowa: CAŁA strona widoczna w panelu.

        Przy takim dopasowaniu ograniczeniem bywa wysokość, więc dla stron
        pionowych poszerzanie okna nie powiększa podglądu — powiększenie
        uzyskuje się wtedy przez Ctrl+kółko albo wyższe okno.
        """
        if not self._has_image:
            return
        self.resetTransform()
        self.fitInView(self._item, Qt.KeepAspectRatio)
        self._zoom = 1.0
        self.verticalScrollBar().setValue(self.verticalScrollBar().minimum())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Przy skali bazowej okno ma nadal mieścić całą stronę; po ręcznym
        # powiększeniu użytkownik sam decyduje o kadrze, więc nie ruszamy.
        if self._has_image and abs(self._zoom - 1.0) < 1e-6:
            self.fit_page()

    def _zoom_by(self, factor):
        new_zoom = min(max(self._zoom * factor, PREVIEW_ZOOM_MIN), PREVIEW_ZOOM_MAX)
        if abs(new_zoom - self._zoom) < 1e-6:
            return
        if abs(new_zoom - PREVIEW_ZOOM_MIN) < 1e-6:
            self.fit_page()
            return
        ratio = new_zoom / self._zoom
        self.scale(ratio, ratio)
        self._zoom = new_zoom

    def wheelEvent(self, event):
        if not self._has_image:
            return
        delta = event.angleDelta().y()

        # Kółko ZAWSZE powiększa. Wcześniej przy dopasowanej stronie pierwszy
        # obrót powiększał, a kolejne — gdy pojawił się już pasek przewijania —
        # zaczynały przesuwać widok; jeden gest robił dwie różne rzeczy.
        # Przesuwanie zostaje na przeciąganiu myszą (ScrollHandDrag).
        forward = (delta > 0) != self._wheel_inverted

        if event.modifiers() & Qt.ShiftModifier and self._multipage:
            self.wheelPaged.emit(-1 if forward else 1)
            return

        self._zoom_by(1.15 if forward else 1 / 1.15)

    def mouseDoubleClickEvent(self, event):
        self.reset_view()


class DragAwareDelegate(TreeItemDelegate):
    """Delegat, który w trakcie przeciągania nie podświetla wiersza pod kursorem.

    Delegat Fluenta maluje szare tło dla KAŻDEGO wiersza ze stanem MouseOver,
    więc podczas przeciągania szarzył się wiersz, nad którym akurat był kursor.
    Miejsce wstawienia pokazuje kreska, a podświetlenie ma dotyczyć wyłącznie
    przenoszonego pliku — czyli stanu Selected, którego tu nie ruszamy.
    """

    def paint(self, painter, option, index):
        if self.parent().drag_active() and (option.state & QStyle.State_MouseOver):
            option.state &= ~QStyle.State_MouseOver
        super().paint(painter, option, index)


class ProjectTree(TreeWidget):
    """Drzewo projektu z przeciąganiem plików (wewnątrz i z Eksploratora).

    Wskaźnik miejsca wstawienia i autoscroll daje Qt. Sam drop przechwytujemy
    i przeliczamy na zmianę w modelu (self.targets w oknie głównym) zamiast
    pozwolić Qt przestawiać elementy — model jest źródłem prawdy, drzewo tylko
    go odrysowuje.
    """

    filesDropped = Signal(list, object, object)   # ścieżki, t_idx, insert_idx
    itemsMoved = Signal(list, object, object)     # [(t,f)], t_idx docelowy, insert_idx

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction)
        # Wskaźnik rysujemy sami (_drop_rect + paintEvent): wbudowany rysuje
        # styl przez PE_IndicatorItemViewItemDrop, a widok ostylowany arkuszem
        # Fluent go gubi — stąd brak linii mimo poprawnego dropu.
        self.setDropIndicatorShown(False)
        self.setAcceptDrops(True)
        self.setMouseTracking(True)
        self.setAutoScroll(True)
        self._drop_rect = None
        self._drag_active = False
        self.setItemDelegate(DragAwareDelegate(self))

        # Własne autoprzewijanie. Wbudowane rusza o 1 px na tik i rozpędza
        # się zbyt wolno — zmierzone 12 px na sekundę przy liście długiej
        # na 1231 px, czyli praktycznie bezużytecznie.
        self.setAutoScroll(False)
        self._drag_pos = None
        self._scroll_step = 0
        self._scroll_timer = QTimer(self)
        self._scroll_timer.setInterval(DRAG_SCROLL_INTERVAL)
        self._scroll_timer.timeout.connect(self._do_drag_scroll)

    def drag_active(self):
        return self._drag_active

    def _set_drag_active(self, active):
        if self._drag_active == active:
            return
        self._drag_active = active
        if not active:
            self._stop_drag_scroll()
        self.viewport().update()

    def mimeTypes(self):
        # Bez "text/uri-list" model odrzuca przeciąganie plików z Eksploratora
        # (canDropMimeData -> False), a wtedy Qt NIE rysuje linii wstawiania —
        # stąd wrażenie, że wskaźnik pojawia się tylko czasem (przy przenoszeniu
        # wewnątrz drzewa działał, przy plikach z dysku nie).
        return super().mimeTypes() + ["text/uri-list"]

    def startDrag(self, supported_actions):
        # Przeciągać można wyłącznie pliki — zestawy zostają na miejscu.
        items = self.selectedItems()
        if not items or any(i.data(0, ROLE_KIND) != "file" for i in items):
            return

        # Własny QDrag zamiast super().startDrag(), bo tylko tak można podać
        # obrazek przeciąganego elementu — widok ostylowany arkuszem Fluent
        # nie rysuje czytelnej miniatury sam z siebie.
        drag = QDrag(self)
        drag.setMimeData(self.model().mimeData([self.indexFromItem(i) for i in items]))
        pixmap = self._drag_pixmap(items)
        drag.setPixmap(pixmap)
        drag.setHotSpot(QPoint(16, pixmap.height() // 2))
        drag.exec(supported_actions, Qt.MoveAction)
        # Po zakończeniu przeciągania (także po upuszczeniu poza drzewem)
        # wracamy do normalnego podświetlania wiersza pod kursorem.
        self._set_drag_active(False)

    def _drag_pixmap(self, items):
        """Rysuje "pigułkę" z nazwą przeciąganego pliku (i licznikiem, gdy więcej)."""
        name = items[0].data(0, ROLE_NAME) or items[0].text(0).strip()
        if len(items) > 1:
            name = f"{name}   +{len(items) - 1}"

        font = self.font()
        metrics = QFontMetrics(font)
        pad_x, pad_y = 14, 8
        text_w = metrics.horizontalAdvance(name)
        width = min(text_w + pad_x * 2, 420)
        height = metrics.height() + pad_y * 2
        # Skracamy dopiero gdy tekst naprawdę się nie mieści — przy równej
        # szerokości elidedText() ucinał go przez zaokrąglenie pomiaru.
        avail = width - pad_x * 2
        elided = name if text_w <= avail else metrics.elidedText(name, Qt.ElideMiddle, avail)

        dpr = self.devicePixelRatioF()
        pixmap = QPixmap(int(width * dpr), int(height * dpr))
        pixmap.setDevicePixelRatio(dpr)
        pixmap.fill(Qt.transparent)

        dark = isDarkTheme()
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(QPen(themeColor(), 1.2))
        painter.setBrush(QColor(43, 43, 43, 235) if dark else QColor(255, 255, 255, 240))
        painter.drawRoundedRect(QRectF(0.5, 0.5, width - 1, height - 1), 6, 6)
        painter.setPen(QColor("#ffffff") if dark else QColor("#1a1a1a"))
        painter.setFont(font)
        painter.drawText(QRectF(pad_x, 0, width - pad_x * 2, height),
                         Qt.AlignVCenter | Qt.AlignLeft, elided)
        painter.end()
        return pixmap

    def dragEnterEvent(self, event):
        self._set_drag_active(True)
        super().dragEnterEvent(event)
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        super().dragMoveEvent(event)
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        self._drag_pos = event.position().toPoint()
        self._update_drag_scroll(self._drag_pos)
        self._drop_rect = self._drop_position(self._drag_pos)[2]
        self.viewport().update()

    def _update_drag_scroll(self, pos):
        """Ustala kierunek i tempo przewijania z odległości od krawędzi.

        Im głębiej w strefie przy krawędzi, tym szybciej — dzięki temu da się
        zarówno dojechać kawałek, jak i przewinąć długą listę na drugi koniec.
        """
        height = self.viewport().height()
        top_in = DRAG_SCROLL_MARGIN - pos.y()
        bottom_in = DRAG_SCROLL_MARGIN - (height - pos.y())

        if top_in > 0:
            depth, direction = top_in, -1
        elif bottom_in > 0:
            depth, direction = bottom_in, 1
        else:
            self._scroll_step = 0
            self._scroll_timer.stop()
            return

        ratio = min(max(depth / DRAG_SCROLL_MARGIN, 0.0), 1.0)
        step = DRAG_SCROLL_MIN_STEP + ratio * (DRAG_SCROLL_MAX_STEP - DRAG_SCROLL_MIN_STEP)
        self._scroll_step = int(round(step)) * direction
        if not self._scroll_timer.isActive():
            self._scroll_timer.start()

    def _do_drag_scroll(self):
        if not self._drag_active or not self._scroll_step:
            self._scroll_timer.stop()
            return
        bar = self.verticalScrollBar()
        before = bar.value()
        bar.setValue(before + self._scroll_step)
        if bar.value() == before:  # koniec listy — nie ma po co tykać dalej
            self._scroll_timer.stop()
            return
        # Lista pojechała pod nieruchomym kursorem, więc linia wstawienia
        # musi zostać przeliczona dla nowej zawartości pod nim.
        if self._drag_pos is not None:
            self._drop_rect = self._drop_position(self._drag_pos)[2]
        self.viewport().update()

    def _stop_drag_scroll(self):
        self._scroll_step = 0
        self._drag_pos = None
        self._scroll_timer.stop()

    def dragLeaveEvent(self, event):
        self._set_drag_active(False)
        self._drop_rect = None
        self.viewport().update()
        super().dragLeaveEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._drop_rect is None:
            return
        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self._drop_rect, themeColor())
        painter.end()

    def _drop_position(self, pos):
        """Zwraca (t_idx, insert_idx, prostokąt linii) dla pozycji kursora.

        Jedno źródło prawdy dla rysowania wskaźnika ORAZ dla samego dropu —
        dzięki temu plik ląduje dokładnie tam, gdzie widać linię. Pozycję
        liczymy z współrzędnych kursora, a nie z dropIndicatorPosition(),
        które zależy od stanu wewnętrznego widoku.
        """
        item = self.itemAt(pos)
        if item is None:
            return None, None, None

        kind = item.data(0, ROLE_KIND)
        if kind == "root":
            return None, None, None

        rect = self.visualItemRect(item)
        left, width = rect.left(), rect.width()

        if kind == "target":
            # Linia tuż pod belką zestawu = wstaw na jego początek.
            t_idx = item.data(0, ROLE_TARGET)
            has_files = item.childCount() > 0
            line = QRect(left, rect.bottom() - 1, width, 2)
            return t_idx, (0 if has_files else None), line

        t_idx = item.data(0, ROLE_TARGET)
        f_idx = item.data(0, ROLE_FILE)
        if pos.y() >= rect.center().y():
            return t_idx, f_idx + 1, QRect(left, rect.bottom() - 1, width, 2)
        return t_idx, f_idx, QRect(left, rect.top() - 1, width, 2)

    def dropEvent(self, event):
        t_idx, insert_idx, _line = self._drop_position(event.position().toPoint())
        self._set_drag_active(False)
        self._drop_rect = None
        self.viewport().update()

        if event.mimeData().hasUrls():
            paths = [u.toLocalFile() for u in event.mimeData().urls()]
            pdfs = [p for p in paths if p.lower().endswith(".pdf")]
            if pdfs:
                self.filesDropped.emit(pdfs, t_idx, insert_idx)
            event.acceptProposedAction()
            return

        if t_idx is None:
            event.ignore()
            return

        moved = [(i.data(0, ROLE_TARGET), i.data(0, ROLE_FILE))
                 for i in self.selectedItems() if i.data(0, ROLE_KIND) == "file"]
        if moved:
            self.itemsMoved.emit(moved, t_idx, insert_idx)
        event.accept()


class MergeWorker(QThread):
    """Scalanie w osobnym wątku — GUI nie zamarza przy dużych plikach."""

    progress = Signal(int, int)
    finished_merge = Signal(int, int, list)

    def __init__(self, targets, output_dir, no_files_template, page_numbers=None):
        super().__init__()
        self.targets = targets
        self.output_dir = output_dir
        self.no_files_template = no_files_template
        # None = bez numeracji; inaczej słownik ustawień z okna konfiguracji.
        self.page_numbers = page_numbers if (page_numbers or {}).get("enabled") else None

    def run(self):
        total = len(self.targets)
        success = 0
        errors = []

        for i, target in enumerate(self.targets, start=1):
            files = target.get("files", [])
            filename = target.get("name", "")
            if not filename.lower().endswith(".pdf"):
                filename += ".pdf"
            out_path = os.path.join(self.output_dir, filename)

            if not files:
                errors.append(self.no_files_template.format(name=target["name"]))
            else:
                try:
                    merger = PdfWriter()
                    for entry in files:
                        fpath = entry.get("path", "") if isinstance(entry, dict) else str(entry)
                        if os.path.exists(fpath):
                            # Zakładka o nazwie pliku źródłowego: w scalonym
                            # zestawie widać, gdzie zaczyna się który rysunek.
                            # Jeśli plik NIESIE JUŻ własną zakładkę o tej samej
                            # nazwie (typowy eksport z CAD-a), własnej nie
                            # dokładamy — inaczej ta sama nazwa pojawia się
                            # dwa razy, jedna pod drugą.
                            name = os.path.splitext(os.path.basename(fpath))[0]
                            if self._has_own_outline_named(fpath, name):
                                merger.append(fpath)
                            else:
                                merger.append(fpath, outline_item=name)
                    merger.write(out_path)
                    merger.close()
                    if self.page_numbers:
                        self._number_pages(out_path)
                    success += 1
                except Exception as exc:
                    errors.append(f"- {target['name']}: {exc}")
                    if os.path.exists(out_path):
                        try:
                            os.remove(out_path)
                        except OSError:
                            pass

            self.progress.emit(i, total)

        self.finished_merge.emit(success, total, errors)

    @staticmethod
    def _normalized(text):
        """Porównanie nazw odporne na spacje, wielkość liter i ".pdf"."""
        text = " ".join((text or "").split()).casefold()
        return text[:-4] if text.endswith(".pdf") else text

    def _has_own_outline_named(self, fpath, name):
        """Czy plik ma JEDNĄ własną zakładkę główną o nazwie pliku?"""
        try:
            doc = fitz.open(fpath)
            roots = [title for level, title, _page in doc.get_toc(simple=True)
                     if level == 1]
            doc.close()
        except Exception:
            return False
        return len(roots) == 1 and self._normalized(roots[0]) == self._normalized(name)

    def _number_pages(self, path):
        """Dopisuje numerację "strona/liczba stron" wg ustawień użytkownika.

        Robi to PyMuPDF już po scaleniu, bo pypdf nie rysuje tekstu. Punkt
        wstawienia przeliczamy macierzą derotacji i podajemy obrót strony —
        inaczej na stronach obróconych (rysunki poziome) numer lądowałby bokiem
        albo poza kartką.
        """
        cfg = self.page_numbers
        position = cfg.get("position", PAGE_NUMBER_DEFAULTS["position"])
        first_page = cfg.get("first_page", True)
        prefix = cfg.get("prefix", "")
        font_size = cfg.get("font_size", PAGE_NUMBER_DEFAULTS["font_size"])
        # Milimetry z okna -> punkty, w których liczy wnętrze PDF-a.
        offset_x = cfg.get("offset_x_mm", PAGE_NUMBER_DEFAULTS["offset_x_mm"]) * MM_TO_POINTS
        offset_y = cfg.get("offset_y_mm", PAGE_NUMBER_DEFAULTS["offset_y_mm"]) * MM_TO_POINTS

        doc = fitz.open(path)
        try:
            total_pages = doc.page_count
            for idx, page in enumerate(doc, start=1):
                if idx == 1 and not first_page:
                    continue
                text = page_number_text(prefix, idx, total_pages)
                rect = page.rect  # rozmiar widoczny, już po obrocie
                width = fitz.get_text_length(text, fontname="helv",
                                             fontsize=font_size)
                if position == "left":
                    x = offset_x
                elif position == "right":
                    x = rect.width - width - offset_x
                else:
                    x = (rect.width - width) / 2 + offset_x
                # Numer ma zostać na kartce nawet przy skrajnym odsunięciu.
                x = min(max(x, 0.0), max(rect.width - width, 0.0))
                y = min(max(rect.height - offset_y, font_size), rect.height)
                page.insert_text(fitz.Point(x, y) * page.derotation_matrix, text,
                                 fontname="helv", fontsize=font_size,
                                 rotate=page.rotation, color=(0, 0, 0))
            doc.save(path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        finally:
            doc.close()


# --- DIALOGI ---
# Wszystkie na MessageBoxBase, który wymaga rodzica (to nakładka na okno).


def _unify_dialog_buttons(dialog):
    """Ujednolica przyciski okna: język i rozmiar czcionki.

    MessageBoxBase zostawia własne napisy ("OK", "Cancel") i nadaje przyciskowi
    anulowania czcionkę w PUNKTACH (9 pt), podczas gdy cały interfejs Fluent
    liczy rozmiar w PIKSELACH (14 px) — przez to "Cancel" wychodził mniejszy od
    wszystkiego wokół i zostawał po angielsku. Tłumaczenie bierzemy z okna
    głównego, żeby nie zmieniać sygnatur wszystkich okien.
    """
    dialog.cancelButton.setFont(dialog.yesButton.font())
    tr = getattr(dialog.parent(), "tr_", None)
    if callable(tr):
        dialog.yesButton.setText(tr("btn_ok"))
        dialog.cancelButton.setText(tr("btn_cancel"))
    # Kolejność przycisków: anulowanie po lewej, zatwierdzenie po prawej.
    # Po przeniesieniu trzeba przywrócić OBU jednakowy udział w szerokości —
    # widżet dodany na nowo dostaje udział 0 i przycisk robi się wąski.
    dialog.buttonLayout.removeWidget(dialog.yesButton)
    dialog.buttonLayout.addWidget(dialog.yesButton, 1)
    for index in range(dialog.buttonLayout.count()):
        dialog.buttonLayout.setStretch(index, 1)

class TextInputDialog(MessageBoxBase):
    def __init__(self, parent, title, label, value="", placeholder=""):
        super().__init__(parent)
        self.viewLayout.addWidget(SubtitleLabel(title, self))
        if label:
            self.viewLayout.addWidget(BodyLabel(label, self))
        self.edit = LineEdit(self)
        self.edit.setText(value)
        self.edit.setPlaceholderText(placeholder)
        self.edit.setClearButtonEnabled(True)
        self.viewLayout.addWidget(self.edit)
        self.widget.setMinimumWidth(460)
        _unify_dialog_buttons(self)
        self.edit.setFocus()
        self.edit.selectAll()

    def value(self):
        return self.edit.text().strip()


class OutputDirDialog(MessageBoxBase):
    def __init__(self, parent, title, label, value, browse_text, browse_title):
        super().__init__(parent)
        self._browse_title = browse_title
        self.viewLayout.addWidget(SubtitleLabel(title, self))
        self.viewLayout.addWidget(BodyLabel(label, self))

        row = QHBoxLayout()
        self.edit = LineEdit(self)
        self.edit.setText(value)
        row.addWidget(self.edit, 1)
        browse = PushButton(browse_text, self, FluentIcon.FOLDER)
        browse.clicked.connect(self._browse)
        row.addWidget(browse)
        self.viewLayout.addLayout(row)
        self.widget.setMinimumWidth(620)
        _unify_dialog_buttons(self)
        self.edit.setFocus()

    def _browse(self):
        # Startuj od katalogu NADRZĘDNEGO wpisanej ścieżki (dla "C:\P\aaa"
        # wejdź do "C:\P"), a gdy i on nie istnieje — szukaj wyżej.
        start = os.path.dirname(self.edit.text().strip())
        while start and not os.path.isdir(start):
            parent = os.path.dirname(start)
            start = parent if parent != start else ""
        picked = QFileDialog.getExistingDirectory(self, self._browse_title, start)
        if picked:
            self.edit.setText(os.path.normpath(picked))

    def value(self):
        return self.edit.text().strip()


class InfoDialog(MessageBoxBase):
    """Proste okno informacyjne/potwierdzenia (tytuł + treść)."""

    def __init__(self, parent, title, body, ok_text="OK", cancel_text=None):
        super().__init__(parent)
        self.viewLayout.addWidget(SubtitleLabel(title, self))
        label = BodyLabel(body, self)
        label.setWordWrap(True)
        self.viewLayout.addWidget(label)
        # Najpierw wspólne ustawienia, potem napisy podane przez wywołującego.
        _unify_dialog_buttons(self)
        self.yesButton.setText(ok_text)
        if cancel_text is None:
            self.cancelButton.hide()
            self.buttonLayout.insertStretch(0, 1)
        else:
            self.cancelButton.setText(cancel_text)
        self.widget.setMinimumWidth(460)


class FileInfoDialog(MessageBoxBase):
    def __init__(self, parent, tr, rows, status_icon, status_color, status_title,
                 status_detail, fpath):
        super().__init__(parent)
        self.viewLayout.addWidget(SubtitleLabel(tr("file_info_dialog_title"), self))

        for label, value in rows:
            row = QHBoxLayout()
            row.addWidget(StrongBodyLabel(label, self))
            val = BodyLabel(value, self)
            val.setWordWrap(True)
            row.addWidget(val, 1)
            self.viewLayout.addLayout(row)

        status_row = QHBoxLayout()
        status_row.addWidget(StrongBodyLabel(tr("row_status"), self))
        marker = StrongBodyLabel(f"{status_icon}  {status_title}", self)
        marker.setTextColor(QColor(status_color), QColor(status_color))
        marker.setWordWrap(True)
        status_row.addWidget(marker, 1)
        self.viewLayout.addLayout(status_row)

        if status_detail:
            detail = CaptionLabel(status_detail, self)
            detail.setWordWrap(True)
            self.viewLayout.addWidget(detail)

        self.viewLayout.addWidget(StrongBodyLabel(tr("row_fullpath"), self))
        path_label = BodyLabel(fpath, self)
        path_label.setWordWrap(True)
        self.viewLayout.addWidget(path_label)

        _unify_dialog_buttons(self)
        self.cancelButton.hide()
        self.buttonLayout.insertStretch(0, 1)
        self.widget.setMinimumWidth(560)


class HelpDialog(MessageBoxBase):
    """Skrócona instrukcja — cztery hasła, bez rozwlekania."""

    def __init__(self, parent, tr):
        super().__init__(parent)
        self.viewLayout.addWidget(SubtitleLabel(tr("help_title"), self))

        for title_key, body_key in (("help_add_t", "help_add"),
                                    ("help_order_t", "help_order"),
                                    ("help_preview_t", "help_preview"),
                                    ("help_dots_t", "help_dots")):
            self.viewLayout.addWidget(StrongBodyLabel(tr(title_key), self))
            body = BodyLabel(tr(body_key), self)
            body.setWordWrap(True)
            self.viewLayout.addWidget(body)

        _unify_dialog_buttons(self)
        self.cancelButton.hide()
        self.buttonLayout.insertStretch(0, 1)
        self.widget.setMinimumWidth(520)


class PageNumberDialog(MessageBoxBase):
    """Konfiguracja numeracji stron scalonego PDF."""

    def __init__(self, parent, tr, settings):
        super().__init__(parent)
        self.tr_ = tr
        self.viewLayout.addWidget(SubtitleLabel(tr("pn_title"), self))

        self.chk_enable = CheckBox(tr("pn_enable"), self)
        self.chk_enable.setChecked(settings["enabled"])
        self.viewLayout.addWidget(self.chk_enable)

        self.chk_first = CheckBox(tr("pn_first"), self)
        self.chk_first.setChecked(settings["first_page"])
        self.viewLayout.addWidget(self.chk_first)

        prefix_row = QHBoxLayout()
        prefix_row.addWidget(BodyLabel(tr("pn_prefix"), self))
        self.edit_prefix = LineEdit(self)
        self.edit_prefix.setMaxLength(PAGE_NUMBER_PREFIX_MAX)
        self.edit_prefix.setText(settings.get("prefix", ""))
        self.edit_prefix.setClearButtonEnabled(True)
        prefix_row.addStretch(1)
        prefix_row.addWidget(self.edit_prefix)
        self.viewLayout.addLayout(prefix_row)

        pos_row = QHBoxLayout()
        pos_row.addWidget(BodyLabel(tr("pn_position"), self))
        self.combo_pos = ComboBox(self)
        self.combo_pos.addItems([tr("pn_left"), tr("pn_center"), tr("pn_right")])
        self.combo_pos.setCurrentIndex(PAGE_NUMBER_POSITIONS.index(settings["position"]))
        pos_row.addStretch(1)
        pos_row.addWidget(self.combo_pos)
        self.viewLayout.addLayout(pos_row)

        self.spin_x = self._spin_row(tr("pn_offset_x"), settings["offset_x_mm"],
                                     -PAGE_NUMBER_OFFSET_MAX, PAGE_NUMBER_OFFSET_MAX,
                                     tr("pn_unit_mm"))
        self.spin_y = self._spin_row(tr("pn_offset_y"), settings["offset_y_mm"],
                                     -PAGE_NUMBER_OFFSET_MAX, PAGE_NUMBER_OFFSET_MAX,
                                     tr("pn_unit_mm"))
        self.spin_font = self._spin_row(tr("pn_font"), settings["font_size"],
                                        PAGE_NUMBER_FONT_MIN, PAGE_NUMBER_FONT_MAX,
                                        tr("pn_unit_pt"))

        # Podpowiedzi tym samym rozmiarem co reszta treści okna (14 px) —
        # drugorzędność niesie kolor, nie mniejsza czcionka.
        self.lbl_hint = BodyLabel("", self)
        self.lbl_hint.setWordWrap(True)
        self.lbl_hint.setTextColor(QColor("#606060"), QColor("#a0a0a0"))
        self.viewLayout.addWidget(self.lbl_hint)
        # Przykład pokazuje DOKŁADNIE to, co wyjdzie na stronie — z aktualnym
        # przedrostkiem, więc nie trzeba zgadywać, czy dojdzie spacja.
        self.lbl_format = BodyLabel("", self)
        self.lbl_format.setTextColor(QColor("#606060"), QColor("#a0a0a0"))
        self.viewLayout.addWidget(self.lbl_format)

        self.chk_enable.stateChanged.connect(self._sync)
        self.combo_pos.currentIndexChanged.connect(self._sync)
        self.edit_prefix.textChanged.connect(self._sync)
        self._sync()
        _unify_dialog_buttons(self)
        self.widget.setMinimumWidth(460)

    def _spin_row(self, label, value, minimum, maximum, suffix):
        row = QHBoxLayout()
        row.addWidget(BodyLabel(label, self))
        spin = SpinBox(self)
        spin.setRange(minimum, maximum)
        spin.setSuffix(suffix)
        spin.setValue(value)
        row.addStretch(1)
        row.addWidget(spin)
        self.viewLayout.addLayout(row)
        return spin

    def _sync(self):
        """Wygaszenie pól, gdy numeracja jest wyłączona — nie ma czego ustawiać."""
        enabled = self.chk_enable.isChecked()
        for widget in (self.chk_first, self.combo_pos, self.spin_x, self.spin_y,
                       self.spin_font, self.edit_prefix):
            widget.setEnabled(enabled)
        centered = self.combo_pos.currentIndex() == 1
        self.lbl_hint.setText(self.tr_("pn_hint_center" if centered else "pn_hint_edge"))
        self.lbl_format.setText(self.tr_(
            "pn_format", example=page_number_text(self.edit_prefix.text(), 3, 12)))

    def value(self):
        return {
            "enabled": self.chk_enable.isChecked(),
            "first_page": self.chk_first.isChecked(),
            "position": PAGE_NUMBER_POSITIONS[self.combo_pos.currentIndex()],
            "prefix": self.edit_prefix.text().strip(),
            "offset_x_mm": self.spin_x.value(),
            "offset_y_mm": self.spin_y.value(),
            "font_size": self.spin_font.value(),
        }


class AboutDialog(MessageBoxBase):
    def __init__(self, parent, tr):
        super().__init__(parent)
        self.viewLayout.addWidget(SubtitleLabel(tr("about_app_name"), self))
        self.viewLayout.addWidget(CaptionLabel(tr("about_version", ver=APP_VERSION), self))

        license_label = StrongBodyLabel(tr("about_license"), self)
        license_label.setTextColor(QColor("#1e8449"), QColor("#2ecc71"))
        self.viewLayout.addWidget(license_label)

        self.viewLayout.addWidget(BodyLabel(tr("about_source"), self))

        components = BodyLabel(tr("about_components"), self)
        components.setWordWrap(True)
        components.setTextColor(QColor("#606060"), QColor("#a0a0a0"))
        self.viewLayout.addWidget(components)

        feedback = BodyLabel(tr("about_feedback"), self)
        feedback.setWordWrap(True)
        self.viewLayout.addWidget(feedback)

        site = BodyLabel(f'<a href="https://www.draftcon.pl">{ABOUT_WEBSITE}</a>', self)
        site.setOpenExternalLinks(True)
        self.viewLayout.addWidget(site)

        mail = BodyLabel(f'<a href="mailto:{ABOUT_EMAIL}">{ABOUT_EMAIL}</a>', self)
        mail.setOpenExternalLinks(True)
        self.viewLayout.addWidget(mail)

        self.viewLayout.addWidget(CaptionLabel(tr("about_version_date", date=ABOUT_DATE), self))

        _unify_dialog_buttons(self)
        self.cancelButton.hide()
        self.buttonLayout.insertStretch(0, 1)
        self.widget.setMinimumWidth(480)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.lang = "pl"
        self.output_dir = ""
        self.project_file_path = ""
        self.selected_date = datetime.date.today()
        self.date_suffix_enabled = True

        today_str = self.selected_date.strftime("%Y.%m.%d")
        self.targets = [
            {"name": f"PZT_{today_str}.pdf", "files": []},
            {"name": f"PAB_{today_str}.pdf", "files": []},
            {"name": f"ZL_{today_str}.pdf", "files": []},
        ]

        self._undo_stack = []
        self._redo_stack = []
        self._undo_limit = 50
        self._dirty = False
        self._merging = False
        self._preview_compact = False  # waski panel = przyciski bez napisow
        self._always_on_top = False    # okno nad innymi (przycisk na pasku)
        # Numeracja stron scalonego PDF — komplet ustawień z okna konfiguracji.
        self._page_number_cfg = dict(PAGE_NUMBER_DEFAULTS,
                                     enabled=False, first_page=True)
        self._last_dir = ""            # ostatni katalog z okien wyboru plików
        self._merge_worker = None
        self._last_moved = set()
        # Zaznaczenie do ustawienia po najbliższej przebudowie drzewa —
        # pole, a nie argument refresh_tree(), bo ta jest podpięta pod timer
        # i pod akcję menu, które wołałyby ją z własnymi argumentami.
        self._select_after_refresh = None

        self._preview_cache = {}
        self._preview_page_counts = {}
        self._preview_path = None
        self._preview_page = 0
        self._preview_page_count = 0

        self._build_ui()
        self._retranslate()
        # Ustawia stan pustego podglądu (w tym podpowiedź sterowania na
        # środku) — bez tego panel byłby pusty do pierwszego zaznaczenia,
        # bo show_preview() woła dopiero obsługa zaznaczenia w drzewie.
        self.show_preview(None)
        self._apply_saved_window_state()
        self.refresh_tree()
        self._refresh_project_history()

        # Stan plików na dysku sprawdzamy przy POWROCIE do okna, a nie z
        # zegara: pliki zmienia się w innych programach, więc przełączenie się
        # z powrotem tutaj jest naturalnym momentem na sprawdzenie — i nie
        # przerywa pracy w trakcie.

    # --- TŁUMACZENIA ---
    def tr_(self, key, **kwargs):
        text = TRANSLATIONS[self.lang].get(key, TRANSLATIONS["pl"].get(key, key))
        return text.format(**kwargs) if kwargs else text

    # --- BUDOWA INTERFEJSU ---
    def _build_ui(self):
        self.resize(*DEFAULT_WINDOW_SIZE)
        # Minimum okna ustala PROGRAM, nie zawartosc. Uklad musi zmiescic sie
        # ponizej tej wartosci (patrz minima w pasie akcji i w panelu
        # podgladu), inaczej Qt podnioslby minimum okna i zwezanie by odbijalo.
        self.setMinimumSize(*MIN_WINDOW_SIZE)

        central = QWidget(self)
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(12, 12, 12, 8)
        root_layout.setSpacing(8)

        # --- GÓRNY PASEK: projekt ---
        proj_row = QHBoxLayout()
        self.combo_project = HistoryComboBox(self)
        self.combo_project.aboutToShow.connect(self._refresh_project_history)
        # Ignored w poziomie: o szerokosci decyduje okno, nie nazwa projektu.
        # Przy setMinimumWidth(320) i podpowiedzi z tresci (287 px) sam ten
        # pasek wyznaczal podloge okna na 508 px — zmierzone.
        self.combo_project.setMinimumWidth(150)
        self.combo_project.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.combo_project.currentTextChanged.connect(self._on_history_selected)
        proj_row.addWidget(self.combo_project, 1)

        self.btn_save_proj = ToolButton(FluentIcon.SAVE, self)
        self.btn_save_proj.clicked.connect(self.save_project)
        proj_row.addWidget(self.btn_save_proj)

        # Znaczek niezapisanych zmian — w Qt to gotowy komponent doczepiany do
        # przycisku; w wersji tkinter trzeba było ręcznie pozycjonować kropkę.
        self.save_badge = InfoBadge.attension("", parent=self.btn_save_proj,
                                              target=self.btn_save_proj)
        self.save_badge.hide()

        self.btn_load_proj = ToolButton(FluentIcon.FOLDER, self)
        self.btn_load_proj.clicked.connect(self.load_project)
        proj_row.addWidget(self.btn_load_proj)

        # Zawsze na wierzchu — przelacznik, bo stan musi byc widoczny na pasku
        # (wcisniety = okno nad innymi). Ikona zmienia sie razem ze stanem.
        self.btn_on_top = ToggleToolButton(FluentIcon.PIN, self)
        self.btn_on_top.setChecked(self._always_on_top)
        self.btn_on_top.toggled.connect(self._toggle_always_on_top)
        proj_row.addWidget(self.btn_on_top)

        self.btn_lang = ToolButton(self)
        self.btn_lang.setIcon(_flag_pixmap(self.lang))
        self.btn_lang.setIconSize(QSize(28, 18))
        self.btn_lang.clicked.connect(self.toggle_language)
        proj_row.addWidget(self.btn_lang)
        root_layout.addLayout(proj_row)

        # --- PASEK AKCJI ---
        actions_row = QHBoxLayout()

        def tool(icon, slot):
            btn = ToolButton(icon, self)
            btn.clicked.connect(slot)
            actions_row.addWidget(btn)
            return btn

        self.btn_outdir = tool(FluentIcon.FOLDER_ADD, self.select_output_dir)
        self.btn_new_target = tool(FluentIcon.DOCUMENT, self.add_target)
        self.btn_add_pdf = tool(FluentIcon.ADD, self.add_files_to_selected)
        self.btn_refresh = tool(FluentIcon.SYNC, self.refresh_tree)
        self.btn_delete = tool(FluentIcon.DELETE, self.delete_selected)
        actions_row.addSpacing(12)
        self.btn_undo = tool(FluentIcon.LEFT_ARROW, self.undo)
        self.btn_redo = tool(FluentIcon.RIGHT_ARROW, self.redo)
        self.btn_undo.setEnabled(False)
        self.btn_redo.setEnabled(False)
        actions_row.addStretch(1)

        # Przelacznik w stylu iOS zamiast kwadratu — stan wlaczone/wylaczone
        # widac z daleka. Napisy On/Off puste, bo opis siedzi w dymku i nie
        # moze poszerzac pasa akcji (patrz podloga szerokosci okna).
        self.chk_date = SwitchButton(self)
        self.chk_date.setOnText("")
        self.chk_date.setOffText("")
        self.chk_date.setChecked(self.date_suffix_enabled)
        self.chk_date.checkedChanged.connect(self._toggle_date_suffix)
        self.chk_date.setMinimumWidth(DATE_SWITCH_MIN_WIDTH)
        actions_row.addWidget(self.chk_date)

        # CalendarPicker sam jest przyciskiem, który po kliknięciu od razu
        # rozwija kalendarz — dialog pośredniczący wymuszałby drugie kliknięcie.
        self.date_picker = CalendarPicker(self)
        self.date_picker.setDateFormat("yyyy.MM.dd")
        self._set_picker_date(self.selected_date)
        self.date_picker.dateChanged.connect(self._on_date_changed)
        self.date_picker.setMinimumWidth(DATE_PICKER_MIN_WIDTH)
        actions_row.addWidget(self.date_picker)
        root_layout.addLayout(actions_row)

        # --- PODZIAŁ: drzewo / podgląd ---
        self.splitter = SplitView(Qt.Horizontal, self)
        self.splitter.setHandleWidth(SPLIT_HANDLE_WIDTH)
        self.splitter.setChildrenCollapsible(False)

        self.tree = ProjectTree(self)
        self.tree.itemSelectionChanged.connect(self._on_tree_selection)
        self.tree.itemDoubleClicked.connect(self._on_double_click)
        self.tree.itemEntered.connect(self._on_item_hover)
        self.tree.filesDropped.connect(self._on_files_dropped)
        self.tree.itemsMoved.connect(self._on_items_moved)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._show_context_menu)
        self.splitter.addWidget(self.tree)

        preview_panel = PreviewPanel(self)
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(8, 0, 0, 0)
        preview_layout.setSpacing(6)

        self.lbl_preview_name = ElidedLabel(self)
        preview_layout.addWidget(self.lbl_preview_name)

        self.preview = PreviewView(self)
        self.preview.wheelPaged.connect(self._change_page)
        # Prawy przycisk w podglądzie otwiera TO SAMO menu co w drzewie.
        self.preview.setContextMenuPolicy(Qt.CustomContextMenu)
        self.preview.customContextMenuRequested.connect(self._show_preview_context_menu)
        preview_layout.addWidget(self.preview, 1)

        # Pasek stron — widoczny tylko dla plików wielostronicowych
        self.page_row = QWidget(self)
        page_layout = QHBoxLayout(self.page_row)
        page_layout.setContentsMargins(0, 0, 0, 0)
        self.btn_prev_page = TransparentToolButton(FluentIcon.LEFT_ARROW, self)
        self.btn_prev_page.clicked.connect(lambda: self._change_page(-1))
        page_layout.addWidget(self.btn_prev_page)
        self.lbl_page = StrongBodyLabel(self)
        self.lbl_page.setAlignment(Qt.AlignCenter)
        page_layout.addWidget(self.lbl_page, 1)
        self.btn_next_page = TransparentToolButton(FluentIcon.RIGHT_ARROW, self)
        self.btn_next_page.clicked.connect(lambda: self._change_page(1))
        page_layout.addWidget(self.btn_next_page)
        # Pasek stron zajmuje miejsce ZAWSZE — chowamy tylko jego zawartosc.
        # Gdy calym paskiem sterowal hide()/show(), wybranie pliku
        # wielostronicowego podnosilo minimum panelu o 36 px (zmierzone)
        # i przy niskim oknie potrafilo je rozciagnac w pionie.
        self.page_row.setFixedHeight(self.page_row.sizeHint().height())
        self._set_page_row_visible(False)
        preview_layout.addWidget(self.page_row)

        self.progress = ProgressBar(self)
        self.progress.setValue(0)
        preview_layout.addWidget(self.progress)

        btn_row = QHBoxLayout()
        # Tekst ustawia _retranslate(), więc tu forma (icon, text, parent).
        self.btn_open_folder = PushButton(FluentIcon.FOLDER, "", self)
        self.btn_open_folder.clicked.connect(self.open_output_folder)
        btn_row.addWidget(self.btn_open_folder, 1)
        self.btn_merge = PrimaryPushButton(FluentIcon.SYNC, "", self)
        self.btn_merge.clicked.connect(self.merge_all_targets)
        btn_row.addWidget(self.btn_merge, 1)
        # Ignored w poziomie: napisy przyciskow (135 + 127 px) nie moga wyznaczac
        # podlogi panelu — to o nia zatrzymywala sie prawa krawedz okna i podglad
        # przestawal sie zwezac. O czytelnosc dba _adapt_preview_buttons(), ktore
        # ponizej PREVIEW_COMPACT_WIDTH zostawia same ikony z dymkiem.
        for _btn in (self.btn_open_folder, self.btn_merge):
            _btn.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
            _btn.setMinimumWidth(40)
        preview_layout.addLayout(btn_row)

        preview_panel.resized.connect(self._adapt_preview_buttons)
        self.splitter.addWidget(preview_panel)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)  # rozszerza się panel podglądu
        self.splitter.setSizes(list(DEFAULT_SPLIT))
        root_layout.addWidget(self.splitter, 1)

        # --- STOPKA ---
        # CaptionLabel, nie BodyLabel — stopka ma mieć drobną czcionkę jak
        # standardowy pasek stanu Windows, a nie rozmiar tekstu treści.
        status = QStatusBar(self)
        status.setSizeGripEnabled(True)
        self.setStatusBar(status)
        self.lbl_status = CaptionLabel(self)
        status.addWidget(self.lbl_status, 1)
        site = CaptionLabel(f'<a href="https://www.draftcon.pl">{ABOUT_WEBSITE}</a>', self)
        site.setOpenExternalLinks(True)
        status.addPermanentWidget(site)

        # Skróty klawiszowe
        for key, slot in (("Ctrl+Z", self.undo), ("Ctrl+Y", self.redo),
                          ("Ctrl+S", self.save_project), ("Del", self.delete_selected),
                          ("F2", self.rename_selected_target)):
            action = QAction(self)
            action.setShortcut(key)
            action.triggered.connect(slot)
            self.addAction(action)

    def _retranslate(self):
        self.setWindowTitle(self.tr_("app_title", ver=APP_VERSION))

        tips = [
            (self.btn_save_proj, "tt_save_project"), (self.btn_load_proj, "tt_open_project"),
            (self.btn_lang, "tt_lang"), (self.btn_on_top, "tt_on_top"),
            (self.btn_outdir, "tt_output_dir"),
            (self.btn_new_target, "tt_new_target"), (self.btn_add_pdf, "tt_add_pdf"),
            (self.btn_refresh, "tt_refresh"), (self.btn_delete, "tt_delete"),
            (self.btn_undo, "tt_undo"), (self.btn_redo, "tt_redo"),
            (self.btn_prev_page, "tt_prev_page"), (self.btn_next_page, "tt_next_page"),
        ]
        for widget, key in tips:
            widget.setToolTip(self.tr_(key))
            widget.installEventFilter(ToolTipFilter(widget, 400, ToolTipPosition.BOTTOM))

        self._apply_date_check_text()
        self._apply_on_top_look()
        self._apply_preview_button_text()
        self.lbl_status.setText(self.tr_("status_ready"))

        # Napis w polu projektu stawia wylacznie _refresh_project_history():
        # ComboBox Fluenta IGNORUJE setCurrentText() dla tekstu spoza listy,
        # wiec pole musi miec wlasna pozycje "(Niezapisany)" na indeksie 0.
        self._refresh_project_history()

        if self._preview_path:
            self.lbl_preview_name.setText(self.tr_("preview_name",
                                                   name=os.path.basename(self._preview_path)))
        else:
            self.lbl_preview_name.setText("")
            # Podpowiedź na środku pustego panelu też musi zmienić język.
            self.preview.set_placeholder(self.tr_("preview_hint"))

        self.btn_save_proj.setToolTip(
            self.tr_("tt_unsaved") if self._dirty else self.tr_("tt_save_project"))

    def toggle_language(self):
        self.lang = "en" if self.lang == "pl" else "pl"
        self.btn_lang.setIcon(_flag_pixmap(self.lang))
        self._retranslate()
        self.refresh_tree()

    # --- UNDO / REDO ---
    def _push_undo(self):
        self._undo_stack.append(copy.deepcopy(self.targets))
        if len(self._undo_stack) > self._undo_limit:
            self._undo_stack.pop(0)
        self._redo_stack.clear()
        self._update_undo_buttons()
        self._mark_dirty()

    def undo(self):
        if not self._undo_stack:
            return
        self._redo_stack.append(copy.deepcopy(self.targets))
        self.targets = self._undo_stack.pop()
        self._last_moved.clear()
        self.refresh_tree()
        self._update_undo_buttons()
        self._mark_dirty()

    def redo(self):
        if not self._redo_stack:
            return
        self._undo_stack.append(copy.deepcopy(self.targets))
        self.targets = self._redo_stack.pop()
        self._last_moved.clear()
        self.refresh_tree()
        self._update_undo_buttons()
        self._mark_dirty()

    def _update_undo_buttons(self):
        self.btn_undo.setEnabled(bool(self._undo_stack))
        self.btn_redo.setEnabled(bool(self._redo_stack))

    # --- WSKAŹNIK NIEZAPISANYCH ZMIAN ---
    def _mark_dirty(self):
        self._dirty = True
        self.save_badge.show()
        self.btn_save_proj.setToolTip(self.tr_("tt_unsaved"))

    def _mark_clean(self):
        self._dirty = False
        self.save_badge.hide()
        self.btn_save_proj.setToolTip(self.tr_("tt_save_project"))

    # --- DRZEWO ---
    def _parse_entry(self, entry):
        if isinstance(entry, dict):
            return entry.get("path", ""), entry.get("mtime")
        return str(entry), None

    def refresh_tree(self):
        expanded = {}
        root = self.tree.topLevelItem(0)
        if root is not None:
            expanded["root"] = root.isExpanded()
            for i in range(root.childCount()):
                expanded[i] = root.child(i).isExpanded()

        # Pozycja przewinięcia musi przeżyć przebudowę: setCurrentItem()
        # przy włączonym autoscrollu przewija listę do kotwicy, więc bez
        # tego odświeżanie co 5 s skakało do zaznaczonego pliku.
        scroll_v = self.tree.verticalScrollBar().value()
        scroll_h = self.tree.horizontalScrollBar().value()

        pending = self._select_after_refresh
        self._select_after_refresh = None

        if pending:
            # Po przeniesieniu pliki mają NOWE indeksy, więc odtwarzanie
            # zaznaczenia po starych zostawiało podświetlenie na przypadkowej
            # pozycji. Tutaj zaznaczenie jedzie razem z plikami.
            selected = set(pending)
            first = min(pending)
            current_key = ("file", first[0], first[1])
        else:
            selected = {(i.data(0, ROLE_TARGET), i.data(0, ROLE_FILE))
                        for i in self.tree.selectedItems()
                        if i.data(0, ROLE_KIND) == "file"}

            # Element BIEŻĄCY to kotwica zaznaczania z Shiftem. clear() go
            # kasuje, a Qt bez kotwicy rozciąga zaznaczenie od pierwszego
            # wiersza — stąd "Shift zaznacza wszystko od góry" po odświeżeniu
            # drzewa (timer co 5 s).
            cur = self.tree.currentItem()
            current_key = None
            if cur is not None:
                current_key = (cur.data(0, ROLE_KIND), cur.data(0, ROLE_TARGET),
                               cur.data(0, ROLE_FILE))
        current_item = None

        self.tree.blockSignals(True)
        self.tree.clear()

        out_dir = self.output_dir or self.tr_("root_not_set")
        root_item = QTreeWidgetItem([self.tr_("root_label", path=out_dir)])
        root_item.setData(0, ROLE_KIND, "root")
        root_item.setIcon(0, FluentIcon.FOLDER.icon())
        font = root_item.font(0)
        font.setBold(True)
        root_item.setFont(0, font)
        root_item.setFlags(root_item.flags() & ~Qt.ItemIsDragEnabled)
        self.tree.addTopLevelItem(root_item)

        for t_idx, target in enumerate(self.targets):
            t_item = QTreeWidgetItem([target["name"]])
            t_item.setData(0, ROLE_KIND, "target")
            t_item.setData(0, ROLE_TARGET, t_idx)
            t_item.setIcon(0, FluentIcon.DOCUMENT.icon())
            t_font = t_item.font(0)
            t_font.setBold(True)
            t_item.setFont(0, t_font)
            t_item.setFlags((t_item.flags() | Qt.ItemIsDropEnabled) & ~Qt.ItemIsDragEnabled)
            root_item.addChild(t_item)
            if current_key == ("target", t_idx, None):
                current_item = t_item

            for f_idx, entry in enumerate(target["files"]):
                fpath, saved_mtime = self._parse_entry(entry)
                fname = os.path.basename(fpath)
                f_item = QTreeWidgetItem([f"{f_idx + 1:>3}.  {fname}"])
                f_item.setData(0, ROLE_KIND, "file")
                f_item.setData(0, ROLE_TARGET, t_idx)
                f_item.setData(0, ROLE_FILE, f_idx)
                f_item.setData(0, ROLE_NAME, fname)
                f_item.setFlags((f_item.flags() | Qt.ItemIsDragEnabled) & ~Qt.ItemIsDropEnabled)

                f_item.setIcon(0, self._dot_for(fpath, saved_mtime, t_idx, f_idx))

                t_item.addChild(f_item)
                if (t_idx, f_idx) in selected:
                    f_item.setSelected(True)
                if current_key == ("file", t_idx, f_idx):
                    current_item = f_item

        root_item.setExpanded(expanded.get("root", True))
        for i in range(root_item.childCount()):
            root_item.child(i).setExpanded(expanded.get(i, True))

        if current_item is None and current_key is not None                 and current_key[0] == "root":
            current_item = root_item
        if current_item is not None:
            # NoUpdate: przywracamy samą kotwicę, bez ruszania zaznaczenia,
            # które odtworzyliśmy wyżej.
            self.tree.setCurrentItem(current_item, 0, QItemSelectionModel.NoUpdate)

        # Lista ZAWSZE zostaje tam, gdzie była — także po przeniesieniu pliku.
        # Przewijanie do przeniesionego elementu wyrzucało widok w inne miejsce
        # (przy pracy na dole listy — na jej górę), a operacja była wykonywana
        # na pozycjach, które użytkownik miał właśnie przed oczami.
        self.tree.verticalScrollBar().setValue(scroll_v)
        self.tree.horizontalScrollBar().setValue(scroll_h)

        self.tree.blockSignals(False)

    def _dot_for(self, fpath, saved_mtime, t_idx, f_idx):
        """Ikona kropki dla pliku.

        Stan niesie kropka przed nazwą, a nie kolor tekstu: kolor = aktualność
        pliku, wypełnienie = element ostatniej operacji (pozostałe mają obrys).
        """
        dark = isDarkTheme()
        # Kolory rozstrzelone celowo: groszkowa zieleń, pomarańcz jak owoc
        # i mocna czerwień. Wcześniejsze (ciemna zieleń, brązowawy pomarańcz,
        # przygaszona ceglasta czerwień) były na małej kropce trudne do
        # odróżnienia od siebie.
        if not os.path.exists(fpath):
            color = "#ff4d4d" if dark else "#e01b1b"
        else:
            curr = os.path.getmtime(fpath)
            if saved_mtime is not None and abs(curr - saved_mtime) > 1.0:
                color = "#ffa033" if dark else "#f57c00"
            else:
                color = "#9ad83c" if dark else "#5faa14"
        return _dot_icon(color, (t_idx, f_idx) in self._last_moved)

    def _refresh_status_dots(self):
        """Przemalowuje same kropki, BEZ przebudowy drzewa.

        Przebudowa (refresh_tree) kasuje wszystko, co żyje w drzewie, a nie w
        projekcie: kotwicę zaznaczania, podświetlenie i pozycję przewinięcia.
        Skoro zmienia się wyłącznie stan plików na dysku, wystarczy podmienić
        ikony na istniejących pozycjach.
        """
        root = self.tree.topLevelItem(0)
        if root is None:
            return
        if root.childCount() != len(self.targets):
            self.refresh_tree()  # struktura się rozjechała — pełna przebudowa
            return

        for t_idx in range(root.childCount()):
            t_item = root.child(t_idx)
            files = self.targets[t_idx]["files"]
            if t_item.childCount() != len(files):
                self.refresh_tree()
                return
            for f_idx, entry in enumerate(files):
                fpath, saved_mtime = self._parse_entry(entry)
                t_item.child(f_idx).setIcon(
                    0, self._dot_for(fpath, saved_mtime, t_idx, f_idx))

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QEvent.ActivationChange and self.isActiveWindow():
            self._refresh_status_dots()

    def _selected_target_index(self):
        for item in self.tree.selectedItems():
            kind = item.data(0, ROLE_KIND)
            if kind in ("target", "file"):
                return item.data(0, ROLE_TARGET)
        return 0

    # --- PRZECIĄGANIE ---
    def _on_files_dropped(self, paths, t_idx, insert_idx):
        if t_idx is None:
            t_idx = self._selected_target_index()
        existing = [self._parse_entry(e)[0] for e in self.targets[t_idx]["files"]]
        new_entries = []
        for p in paths:
            p = os.path.normpath(p)
            if p not in existing:
                mtime = os.path.getmtime(p) if os.path.exists(p) else None
                new_entries.append({"path": p, "mtime": mtime})
        if not new_entries:
            return

        self._push_undo()
        files = self.targets[t_idx]["files"]
        start = len(files) if insert_idx is None else insert_idx
        files[start:start] = new_entries
        self._last_moved = {(t_idx, start + i) for i in range(len(new_entries))}
        self.refresh_tree()

    def _on_items_moved(self, moved, dest_t, insert_idx):
        self._push_undo()

        # Indeks docelowy trzeba skorygować o elementy usuwane PRZED nim
        # w tym samym zestawie — inaczej wstawienie trafi o tyle pozycji dalej.
        if insert_idx is not None:
            removed_before = sum(1 for t, f in moved if t == dest_t and f < insert_idx)
            insert_idx -= removed_before

        entries = []
        for t, f in sorted(moved, key=lambda x: (x[0], x[1]), reverse=True):
            if t < len(self.targets) and f < len(self.targets[t]["files"]):
                entries.append(self.targets[t]["files"].pop(f))
        entries.reverse()

        files = self.targets[dest_t]["files"]
        start = len(files) if insert_idx is None else max(0, min(insert_idx, len(files)))
        files[start:start] = entries
        self._last_moved = {(dest_t, start + i) for i in range(len(entries))}
        # Podświetlenie ma pojechać razem z plikami, a nie zostać na starych
        # pozycjach, gdzie po przeniesieniu stoją już inne pliki.
        self._select_after_refresh = set(self._last_moved)
        self.refresh_tree()
        self._on_tree_selection()  # podgląd idzie za nowym zaznaczeniem

    # --- PODGLĄD ---
    def _on_tree_selection(self):
        items = self.tree.selectedItems()
        if len(items) != 1 or items[0].data(0, ROLE_KIND) != "file":
            self.show_preview(None)
            return
        item = items[0]
        t_idx, f_idx = item.data(0, ROLE_TARGET), item.data(0, ROLE_FILE)
        if t_idx >= len(self.targets) or f_idx >= len(self.targets[t_idx]["files"]):
            self.show_preview(None)
            return
        self.show_preview(self._parse_entry(self.targets[t_idx]["files"][f_idx])[0])

    def show_preview(self, fpath, reset_view=True):
        # Widok resetujemy tylko przy ZMIANIE pliku — odświeżanie drzewa co 5 s
        # odtwarza zaznaczenie i bez tego kasowałoby powiększenie.
        changed = fpath != self._preview_path
        self._preview_path = fpath
        if changed:
            self._preview_page = 0

        if not fpath:
            # Bez napisu "nic nie zaznaczono" — zostaje pusta, biała ramka
            # podglądu, a na jej środku skrót sterowania.
            self.lbl_preview_name.setText("")
            self.preview.clear_image()
            self.preview.set_placeholder(self.tr_("preview_hint"))
            self._preview_page_count = 0
            self._update_page_row()
            return

        self.lbl_preview_name.setText(self.tr_("preview_name", name=os.path.basename(fpath)))

        if not os.path.exists(fpath):
            self._preview_page_count = 0
            self.preview.clear_image()
            self.preview.set_placeholder(self.tr_("preview_missing"), error=True)
            self._update_page_row()
            return

        pixmap = self._page_pixmap(fpath, self._preview_page)
        if pixmap is None:
            self._preview_page_count = 0
            self.preview.clear_image()
            self.preview.set_placeholder(self.tr_("preview_error"), error=True)
            self._update_page_row()
            return

        self.preview.set_placeholder("")
        self.preview.set_multipage(self._preview_page_count > 1)
        self.preview.set_pixmap(pixmap, keep_view=not (changed or reset_view))
        self._update_page_row()

    def _page_pixmap(self, fpath, page_no):
        try:
            mtime = os.path.getmtime(fpath)
        except OSError:
            return None

        cached = self._preview_cache.get((fpath, page_no))
        counted = self._preview_page_counts.get(fpath)
        if cached and cached[0] == mtime and counted and counted[0] == mtime:
            self._preview_page_count = counted[1]
            return cached[1]

        try:
            doc = fitz.open(fpath)
            count = doc.page_count
            page_no = min(max(page_no, 0), count - 1)
            page = doc[page_no]
            zoom = PREVIEW_RENDER_WIDTH / page.rect.width
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            # PyMuPDF -> QImage bez pośrednictwa Pillow. copy() jest konieczne:
            # QImage nie przejmuje na własność bufora pix.samples, a pixmapa
            # PyMuPDF zostanie zwolniona po zamknięciu dokumentu.
            image = QImage(pix.samples, pix.width, pix.height, pix.stride,
                           QImage.Format_RGB888).copy()
            doc.close()
        except Exception:
            return None

        pixmap = QPixmap.fromImage(image)
        self._preview_page = page_no
        self._preview_page_count = count
        self._preview_page_counts[fpath] = (mtime, count)
        self._preview_cache[(fpath, page_no)] = (mtime, pixmap)
        while len(self._preview_cache) > PREVIEW_CACHE_LIMIT:
            self._preview_cache.pop(next(iter(self._preview_cache)))
        return pixmap

    def _update_page_row(self):
        if self._preview_page_count > 1:
            self.lbl_page.setText(f"{self._preview_page + 1} / {self._preview_page_count}")
            self.btn_prev_page.setEnabled(self._preview_page > 0)
            self.btn_next_page.setEnabled(self._preview_page < self._preview_page_count - 1)
            self._set_page_row_visible(True)
        else:
            self._set_page_row_visible(False)

    def changeEvent(self, event):
        # Zmiana koloru wyrozniajacego w Windows dochodzi tu jako zmiana
        # palety aplikacji — wtedy odswiezamy kolor motywu i przemalowujemy
        # okno, zeby przelaczniki i przycisk scalania poszly za systemem.
        super().changeEvent(event)
        if event.type() in (QEvent.ApplicationPaletteChange, QEvent.PaletteChange):
            apply_system_accent()
            self.update()

    def _toggle_always_on_top(self, enabled):
        """Okno nad innymi. Na Windows zmiana flagi ukrywa okno, wiec po
        setWindowFlag trzeba je pokazac ponownie — bez tego znika z ekranu."""
        self._always_on_top = bool(enabled)
        visible = self.isVisible()
        self.setWindowFlag(Qt.WindowStaysOnTopHint, self._always_on_top)
        if visible:
            self.show()
        self._apply_on_top_look()

    def _apply_on_top_look(self):
        """Ikona i dymek ida za stanem przelacznika."""
        self.btn_on_top.setIcon(FluentIcon.PIN if self._always_on_top else FluentIcon.UNPIN)
        self.btn_on_top.setToolTip(
            self.tr_("tt_on_top") if self._always_on_top else self.tr_("tt_on_top_off"))

    def _apply_date_check_text(self):
        """Przelacznik daty bez napisu — opis zostaje w dymku."""
        self.chk_date.setToolTip(self.tr_("chk_date_suffix"))

    def _adapt_preview_buttons(self, width):
        compact = width < PREVIEW_COMPACT_WIDTH
        if compact == self._preview_compact:
            return
        self._preview_compact = compact
        self._apply_preview_button_text()

    def _apply_preview_button_text(self):
        """Napisy przyciskow — w waskim panelu zostaja same ikony z dymkiem."""
        merge_key = "btn_merging" if self._merging else "btn_merge"
        for btn, key in ((self.btn_open_folder, "btn_open_folder"),
                         (self.btn_merge, merge_key)):
            full = self.tr_(key)
            btn.setToolTip(full)
            btn.setText("" if self._preview_compact else full)

    def _set_page_row_visible(self, visible):
        """Pokazuje/ukrywa zawartosc paska stron, nie ruszajac jego wysokosci."""
        for w in (self.btn_prev_page, self.lbl_page, self.btn_next_page):
            w.setVisible(visible)

    def _change_page(self, delta):
        if self._preview_page_count <= 1 or not self._preview_path:
            return
        new_page = min(max(self._preview_page + delta, 0), self._preview_page_count - 1)
        if new_page == self._preview_page:
            return
        self._preview_page = new_page
        self.show_preview(self._preview_path, reset_view=False)

    # --- STOPKA / HOVER ---
    def _on_item_hover(self, item, column):
        kind = item.data(0, ROLE_KIND)
        if kind == "root":
            self.lbl_status.setText(self.tr_("status_hover_root",
                                             path=self.output_dir or self.tr_("root_not_set")))
        elif kind == "target":
            t_idx = item.data(0, ROLE_TARGET)
            name = self.targets[t_idx]["name"]
            full = os.path.join(self.output_dir, name) if self.output_dir else name
            self.lbl_status.setText(self.tr_("status_hover_target", path=full))
        elif kind == "file":
            t_idx, f_idx = item.data(0, ROLE_TARGET), item.data(0, ROLE_FILE)
            fpath = self._parse_entry(self.targets[t_idx]["files"][f_idx])[0]
            self.lbl_status.setText(self.tr_("status_hover_file", path=fpath))

    # --- KATALOG DOCELOWY ---
    def process_output_dir_path(self, path):
        if not path:
            return ""
        normalized = os.path.normpath(path)
        if not os.path.exists(normalized):
            try:
                os.makedirs(normalized, exist_ok=True)
            except Exception as exc:
                self._error(self.tr_("create_dir_err_title"),
                            self.tr_("create_dir_err_text", err=str(exc)))
        return normalized

    def select_output_dir(self, on_confirm=None):
        dialog = OutputDirDialog(self, self.tr_("outdir_dialog_title"),
                                 self.tr_("outdir_dialog_label"), self.output_dir,
                                 self.tr_("browse_btn"), self.tr_("browse_dialog_title"))
        if dialog.exec():
            raw = dialog.value()
            if raw:
                self.output_dir = self.process_output_dir_path(raw)
                self._mark_dirty()
                self.refresh_tree()
                if on_confirm:
                    on_confirm()

    def open_output_folder(self):
        if not self.output_dir or not os.path.isdir(self.output_dir):
            self._warn(self.tr_("warn_title"), self.tr_("warn_no_output_dir"))
            return
        os.startfile(self.output_dir)

    # --- DATA ---
    def _strip_date_suffix(self, name):
        return DATE_SUFFIX_RE.sub(r'\2', name)

    def _apply_date_suffix(self, name, date_str):
        base = self._strip_date_suffix(name)
        if base.lower().endswith(".pdf"):
            return f"{base[:-4]}_{date_str}.pdf"
        return f"{base}_{date_str}"

    def _set_picker_date(self, date):
        # setDate() emituje dateChanged, więc przy ustawianiu programowym
        # blokujemy sygnał — inaczej wpadlibyśmy w pętlę obsługi zmiany.
        self.date_picker.blockSignals(True)
        self.date_picker.setDate(QDate(date.year, date.month, date.day))
        self.date_picker.blockSignals(False)

    def _on_date_changed(self, qdate):
        if not qdate.isValid():
            return
        new_date = datetime.date(qdate.year(), qdate.month(), qdate.day())
        if new_date == self.selected_date:
            return

        self.selected_date = new_date
        if self.date_suffix_enabled:
            self._push_undo()
            date_str = new_date.strftime("%Y.%m.%d")
            for target in self.targets:
                target["name"] = self._apply_date_suffix(target["name"], date_str)
            self.refresh_tree()

    def _toggle_date_suffix(self):
        self.date_suffix_enabled = self.chk_date.isChecked()
        self.date_picker.setEnabled(self.date_suffix_enabled)

        self._push_undo()
        date_str = self.selected_date.strftime("%Y.%m.%d")
        for target in self.targets:
            if self.date_suffix_enabled:
                if not DATE_SUFFIX_RE.search(target["name"]):
                    target["name"] = self._apply_date_suffix(target["name"], date_str)
            else:
                target["name"] = self._strip_date_suffix(target["name"])
        self.refresh_tree()

    # --- ZESTAWY I PLIKI ---
    def _name_exists(self, name, exclude_idx=None):
        return any(i != exclude_idx and t["name"].lower() == name.lower()
                   for i, t in enumerate(self.targets))

    def add_target(self):
        dialog = TextInputDialog(self, self.tr_("new_target_dialog_title"),
                                 self.tr_("new_target_dialog_label"))
        if not dialog.exec():
            return
        name = dialog.value()
        if not name:
            return
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        if self._name_exists(name):
            self._warn(self.tr_("warn_title"), self.tr_("name_exists_warn", name=name))
            return
        self._push_undo()
        self.targets.append({"name": name, "files": []})
        self.refresh_tree()

    def rename_selected_target(self):
        items = self.tree.selectedItems()
        if not items:
            return
        item = items[0]
        if item.data(0, ROLE_KIND) != "target":
            return
        t_idx = item.data(0, ROLE_TARGET)
        current = self.targets[t_idx]["name"]

        dialog = TextInputDialog(self, self.tr_("rename_dialog_title"), "", current)
        if not dialog.exec():
            return
        new_name = dialog.value()
        if not new_name:
            return
        if not new_name.lower().endswith(".pdf"):
            new_name += ".pdf"
        if new_name == current:
            return
        if self._name_exists(new_name, exclude_idx=t_idx):
            self._warn(self.tr_("warn_title"), self.tr_("name_exists_warn", name=new_name))
            return
        self._push_undo()
        self.targets[t_idx]["name"] = new_name
        self.refresh_tree()

    def add_files_to_selected(self):
        t_idx = self._selected_target_index()
        insert_idx = None
        for item in self.tree.selectedItems():
            if item.data(0, ROLE_KIND) == "file":
                insert_idx = item.data(0, ROLE_FILE) + 1
                break

        files, _ = QFileDialog.getOpenFileNames(self, self.tr_("ctx_add_pdf"),
                                                self._dialog_dir(),
                                                self.tr_("file_filter_pdf"))
        if files:
            self._remember_dir(files[0])
            self._on_files_dropped([os.path.normpath(f) for f in files], t_idx, insert_idx)

    def delete_selected(self):
        items = self.tree.selectedItems()
        if not items:
            return

        targets_to_delete = set()
        files_to_delete = []
        for item in items:
            kind = item.data(0, ROLE_KIND)
            if kind == "target":
                targets_to_delete.add(item.data(0, ROLE_TARGET))
            elif kind == "file":
                files_to_delete.append((item.data(0, ROLE_TARGET), item.data(0, ROLE_FILE)))

        if targets_to_delete:
            if len(self.targets) - len(targets_to_delete) < 1:
                self._warn(self.tr_("warn_title"), self.tr_("warn_min_target"))
                return
            confirm = InfoDialog(self, self.tr_("confirm_title"),
                                 self.tr_("confirm_delete_targets", n=len(targets_to_delete)),
                                 self.tr_("btn_ok"), self.tr_("btn_cancel"))
            if not confirm.exec():
                return
            self._push_undo()
            for t_idx in sorted(targets_to_delete, reverse=True):
                del self.targets[t_idx]
            self._last_moved.clear()
            self.refresh_tree()
            return

        if files_to_delete:
            self._push_undo()
            for t_idx, f_idx in sorted(files_to_delete, key=lambda x: (x[0], x[1]), reverse=True):
                if t_idx < len(self.targets) and f_idx < len(self.targets[t_idx]["files"]):
                    del self.targets[t_idx]["files"][f_idx]
            self._last_moved.clear()
            self.refresh_tree()

    def sort_selected(self, ascending=True):
        by_branch = {}
        for item in self.tree.selectedItems():
            if item.data(0, ROLE_KIND) != "file":
                continue
            by_branch.setdefault(item.data(0, ROLE_TARGET), []).append(item.data(0, ROLE_FILE))
        if not any(len(v) >= 2 for v in by_branch.values()):
            return

        self._push_undo()
        for t_idx, f_idxs in by_branch.items():
            slots = sorted(f_idxs)
            entries = [self.targets[t_idx]["files"][i] for i in slots]
            entries.sort(key=lambda e: os.path.basename(self._parse_entry(e)[0]).lower(),
                         reverse=not ascending)
            for slot, entry in zip(slots, entries):
                self.targets[t_idx]["files"][slot] = entry
        self.refresh_tree()

    def _on_double_click(self, item, column):
        kind = item.data(0, ROLE_KIND)
        if kind == "root":
            self.select_output_dir()
        elif kind == "target":
            self.rename_selected_target()
        elif kind == "file":
            self.open_selected_file(item.data(0, ROLE_TARGET), item.data(0, ROLE_FILE))

    def _file_path_at(self, t_idx, f_idx):
        if t_idx is None or f_idx is None:
            return None
        if t_idx >= len(self.targets) or f_idx >= len(self.targets[t_idx]["files"]):
            return None
        return self._parse_entry(self.targets[t_idx]["files"][f_idx])[0]

    def open_selected_file(self, t_idx, f_idx):
        fpath = self._file_path_at(t_idx, f_idx)
        if not fpath or not os.path.exists(fpath):
            self._error(self.tr_("err_title"), self.tr_("err_file_missing", path=fpath or ""))
            return
        os.startfile(fpath)

    def reveal_in_explorer(self, t_idx, f_idx):
        fpath = self._file_path_at(t_idx, f_idx)
        if not os.path.exists(fpath):
            self._error(self.tr_("err_title"), self.tr_("err_file_missing", path=fpath))
            return
        # Dokładnie ten format wiersza poleceń: cudzysłów tylko wokół ścieżki,
        # sklejony z /select — lista argumentów psuje ścieżki ze spacjami.
        subprocess.Popen(f'explorer /select,"{os.path.normpath(fpath)}"')

    # --- MENU KONTEKSTOWE ---
    def _append_common_menu(self, menu):
        """Część wspólna menu podręcznego — ta sama w drzewie i w podglądzie."""
        if not menu.isEmpty():
            menu.addSeparator()
        wheel = QAction(self.tr_("menu_invert_wheel"), menu)
        wheel.setCheckable(True)
        wheel.setChecked(self.preview.wheel_inverted())
        wheel.toggled.connect(self.preview.set_wheel_inverted)
        menu.addAction(wheel)

        menu.addAction(self.tr_("menu_page_numbers"), self.configure_page_numbers)

        menu.addSeparator()
        menu.addAction(self.tr_("ctx_help"), self.show_help)
        menu.addAction(self.tr_("ctx_about"), self.show_about)

    def _dialog_dir(self, preferred=""):
        """Katalog startowy okna wyboru plików.

        Pierwszeństwo ma miejsce związane z operacją (np. katalog otwartego
        projektu), a gdy go nie ma — ostatni katalog, z którego użytkownik coś
        wybrał. Pusty ciąg oznacza, że decyduje system.
        """
        if preferred and os.path.isdir(preferred):
            return preferred
        return self._last_dir if os.path.isdir(self._last_dir) else ""

    def _remember_dir(self, path):
        directory = path if os.path.isdir(path) else os.path.dirname(path)
        if os.path.isdir(directory):
            self._last_dir = directory

    def configure_page_numbers(self):
        dialog = PageNumberDialog(self, self.tr_, self._page_number_cfg)
        if dialog.exec():
            self._page_number_cfg = dialog.value()

    def _show_preview_context_menu(self, pos):
        menu = QMenu(self)
        self._append_common_menu(menu)
        menu.exec(self.preview.viewport().mapToGlobal(pos))

    def _show_context_menu(self, pos):
        item = self.tree.itemAt(pos)
        if item is None:
            # Puste miejsce w drzewie — zostaje część wspólna, żeby prawy
            # przycisk zawsze coś dawał.
            menu = QMenu(self)
            self._append_common_menu(menu)
            menu.exec(self.tree.viewport().mapToGlobal(pos))
            return
        if not item.isSelected():
            self.tree.setCurrentItem(item)

        kind = item.data(0, ROLE_KIND)
        menu = QMenu(self)

        if kind == "root":
            menu.addAction(self.tr_("ctx_set_output_dir"), self.select_output_dir)
            menu.addAction(self.tr_("ctx_open_output_folder"), self.open_output_folder)
            menu.addSeparator()
            menu.addAction(self.tr_("ctx_new_target"), self.add_target)
            menu.addAction(self.tr_("ctx_refresh"), self.refresh_tree)
        elif kind == "target":
            menu.addAction(self.tr_("ctx_rename"), self.rename_selected_target)
            menu.addAction(self.tr_("ctx_add_pdf"), self.add_files_to_selected)
            menu.addSeparator()
            menu.addAction(self.tr_("ctx_new_target"), self.add_target)
            menu.addAction(self.tr_("ctx_refresh"), self.refresh_tree)
            menu.addSeparator()
            menu.addAction(self.tr_("ctx_delete_target"), self.delete_selected)
        else:
            # Do domknięć trafiają INDEKSY, nie obiekty QTreeWidgetItem: menu
            # ma własną pętlę zdarzeń, więc timer odświeżający drzewo co 5 s
            # może w jego trakcie wywołać refresh_tree() i skasować obiekty
            # C++ elementów — użycie ich potem rzuca RuntimeError.
            t_idx, f_idx = item.data(0, ROLE_TARGET), item.data(0, ROLE_FILE)
            menu.addAction(self.tr_("ctx_open_file"), lambda: self.open_selected_file(t_idx, f_idx))
            menu.addAction(self.tr_("ctx_reveal"), lambda: self.reveal_in_explorer(t_idx, f_idx))
            menu.addSeparator()
            menu.addAction(self.tr_("ctx_add_pdf"), self.add_files_to_selected)
            menu.addAction(self.tr_("ctx_details"), lambda: self.show_file_info(t_idx, f_idx))
            if len([i for i in self.tree.selectedItems()
                    if i.data(0, ROLE_KIND) == "file"]) >= 2:
                menu.addSeparator()
                menu.addAction(self.tr_("ctx_sort_az"), lambda: self.sort_selected(True))
                menu.addAction(self.tr_("ctx_sort_za"), lambda: self.sort_selected(False))
            menu.addSeparator()
            menu.addAction(self.tr_("ctx_delete_file"), self.delete_selected)

        self._append_common_menu(menu)
        menu.exec(self.tree.viewport().mapToGlobal(pos))

    # --- INFORMACJE O PLIKU ---
    def _page_size_text(self, fpath):
        try:
            doc = fitz.open(fpath)
            rect = doc[0].rect
            pages = doc.page_count
            doc.close()
        except Exception:
            return self.tr_("page_size_error"), None

        w_mm = rect.width / 72 * 25.4
        h_mm = rect.height / 72 * 25.4
        w_sorted, h_sorted = min(w_mm, h_mm), max(w_mm, h_mm)

        name = None
        for pname, pw, ph in PAPER_SIZES_MM:
            if abs(w_sorted - pw) <= 2 and abs(h_sorted - ph) <= 2:
                name = pname
                break

        orientation = self.tr_("orientation_h") if w_mm > h_mm else self.tr_("orientation_v")
        extra = ", ".join(x for x in (name, orientation) if x)
        size_str = f"{w_mm:.0f} × {h_mm:.0f} mm"
        return (f"{size_str} ({extra})" if extra else size_str), pages

    def show_file_info(self, t_idx, f_idx):
        if t_idx >= len(self.targets) or f_idx >= len(self.targets[t_idx]["files"]):
            return
        fpath, saved_mtime = self._parse_entry(self.targets[t_idx]["files"][f_idx])
        fname = os.path.basename(fpath)

        if not os.path.exists(fpath):
            icon, color = "✕", "#e74c3c"
            title = self.tr_("status_missing_title")
            detail = ""
            size_text, pages = "—", None
        else:
            size_text, pages = self._page_size_text(fpath)
            curr = os.path.getmtime(fpath)
            curr_dt = datetime.datetime.fromtimestamp(curr).strftime('%Y-%m-%d %H:%M:%S')
            if saved_mtime is not None and abs(curr - saved_mtime) > 1.0:
                saved_dt = datetime.datetime.fromtimestamp(saved_mtime).strftime('%Y-%m-%d %H:%M:%S')
                icon, color = "⚠", "#e67e22"
                title = self.tr_("status_outdated_title")
                detail = self.tr_("detail_saved_disk", saved=saved_dt, curr=curr_dt)
            else:
                icon, color = "✓", "#27ae60"
                title = self.tr_("status_current_title")
                detail = self.tr_("detail_last_modified", curr=curr_dt)

        rows = [
            (self.tr_("row_index"), str(f_idx + 1)),
            (self.tr_("row_filename"), fname),
            (self.tr_("row_pagesize"), size_text),
        ]
        if pages:
            rows.append((self.tr_("row_pages"), str(pages)))

        FileInfoDialog(self, self.tr_, rows, icon, color, title, detail, fpath).exec()

    def show_help(self):
        HelpDialog(self, self.tr_).exec()

    def show_about(self):
        AboutDialog(self, self.tr_).exec()

    # --- SCALANIE ---
    def merge_all_targets(self):
        if not self.targets:
            self._error(self.tr_("err_title"), self.tr_("err_no_targets"))
            return
        if self._merging:
            return
        if not self.output_dir:
            self.select_output_dir(on_confirm=self.merge_all_targets)
            return

        self.output_dir = self.process_output_dir_path(self.output_dir)

        self._merging = True
        self.btn_merge.setEnabled(False)
        self._apply_preview_button_text()
        self.progress.setValue(0)

        snapshot = [dict(t, files=list(t.get("files", []))) for t in self.targets]
        self._merge_worker = MergeWorker(snapshot, self.output_dir, self.tr_("merge_no_files"),
                                         page_numbers=dict(self._page_number_cfg))
        self._merge_worker.progress.connect(self._on_merge_progress)
        self._merge_worker.finished_merge.connect(self._on_merge_done)
        self._merge_worker.start()

    def _on_merge_progress(self, done, total):
        self.progress.setValue(int(done / total * 100))

    def _on_merge_done(self, success, total, errors):
        self._merging = False
        self.progress.setValue(0)
        self.btn_merge.setEnabled(True)
        self._apply_preview_button_text()
        self.refresh_tree()

        if errors:
            InfoDialog(self, self.tr_("merge_result_title"),
                       self.tr_("merge_result_text", success=success, total=total,
                                errors="\n".join(errors)), self.tr_("btn_ok")).exec()
            return

        self._play_success_sound()
        bar = InfoBar.success(self.tr_("success_title"), self.tr_("success_text"),
                              duration=6000, position=InfoBarPosition.TOP, parent=self)
        button = PushButton(self.tr_("open_folder_btn"), bar, FluentIcon.FOLDER)
        button.clicked.connect(self.open_output_folder)
        bar.addWidget(button)

    def _play_success_sound(self):
        tada = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Media", "tada.wav")
        try:
            if os.path.exists(tada):
                winsound.PlaySound(tada, winsound.SND_FILENAME | winsound.SND_ASYNC)
            else:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass

    # --- PROJEKT: ZAPIS / ODCZYT / HISTORIA ---
    def save_project(self):
        path, _ = QFileDialog.getSaveFileName(
            self, self.tr_("tt_save_project"),
            self.project_file_path or self._dialog_dir(),
            self.tr_("file_filter_project"))
        if not path:
            return
        self._remember_dir(path)

        if self.output_dir:
            self.output_dir = self.process_output_dir_path(self.output_dir)

        for target in self.targets:
            updated = []
            for entry in target["files"]:
                fpath, _unused = self._parse_entry(entry)
                mtime = os.path.getmtime(fpath) if os.path.exists(fpath) else None
                updated.append({"path": fpath, "mtime": mtime})
            target["files"] = updated

        data = {"version": 4.0, "output_dir": self.output_dir, "targets": self.targets}
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as exc:
            self._error(self.tr_("err_title"), str(exc))
            return

        self.project_file_path = path
        self._set_project_label(path)
        self._add_to_history(path)
        self._mark_clean()
        self.refresh_tree()

    def load_project(self):
        path, _ = QFileDialog.getOpenFileName(
            self, self.tr_("tt_open_project"),
            self._dialog_dir(os.path.dirname(self.project_file_path)
                             if self.project_file_path else ""),
            self.tr_("file_filter_project"))
        if path:
            self._remember_dir(path)
            self._load_project_from_path(path)

    def _load_project_from_path(self, path):
        if not os.path.exists(path):
            self._error(self.tr_("err_title"), self.tr_("err_file_missing", path=path))
            self._remove_from_history(path)
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as exc:
            self._error(self.tr_("err_title"), str(exc))
            return

        self.output_dir = data.get("output_dir", "")
        self.targets = data.get("targets", [])
        self.project_file_path = path
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._update_undo_buttons()
        self._last_moved.clear()
        self._set_project_label(path)
        self._add_to_history(path)
        self._mark_clean()
        self.refresh_tree()

    def _set_project_label(self, path):
        self.combo_project.blockSignals(True)
        self._select_project_item(path)
        self.combo_project.blockSignals(False)

    # --- ZAPAMIĘTANY ROZMIAR OKNA I POZYCJA SUWAKA ---
    def _sanitize_page_numbers(self, data):
        """Ustawienia numeracji z pliku — każde pole sprawdzane osobno.

        Plik stanu może pochodzić ze starszej wersji albo zostać ręcznie
        zepsuty; wtedy pojedyncze pole wraca do wartości domyślnej, zamiast
        wywracać wczytywanie całego stanu okna.
        """
        cfg = dict(PAGE_NUMBER_DEFAULTS, enabled=False, first_page=True)
        if not isinstance(data, dict):
            return cfg
        # "enabled" NIE wraca z pliku — numeracja startuje wylaczona przy
        # kazdym uruchomieniu, zeby nie dopisac numerow w scaleniu, o ktorym
        # nikt w tej sesji nie decydowal. Reszta ustawien (pozycja, prefiks,
        # czcionka, odsuniecia) zostaje zapamietana.
        if isinstance(data.get("first_page"), bool):
            cfg["first_page"] = data["first_page"]
        if data.get("position") in PAGE_NUMBER_POSITIONS:
            cfg["position"] = data["position"]
        for key in ("offset_x_mm", "offset_y_mm"):
            value = data.get(key)
            if isinstance(value, int) and abs(value) <= PAGE_NUMBER_OFFSET_MAX:
                cfg[key] = value
        size = data.get("font_size")
        if isinstance(size, int) and PAGE_NUMBER_FONT_MIN <= size <= PAGE_NUMBER_FONT_MAX:
            cfg["font_size"] = size
        prefix = data.get("prefix")
        if isinstance(prefix, str):
            cfg["prefix"] = prefix.strip()[:PAGE_NUMBER_PREFIX_MAX]
        return cfg

    def _window_state_file(self):
        return os.path.join(get_app_dir(), WINDOW_STATE_FILENAME)

    def _apply_saved_window_state(self):
        try:
            with open(self._window_state_file(), "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return  # brak pliku albo uszkodzony — zostają wartości domyślne

        size = data.get("window")
        if isinstance(size, list) and len(size) == 2 and all(isinstance(v, int) for v in size):
            w, h = size
            # Przycięcie do ekranu: rozmiar zapisany na większym monitorze nie
            # może otworzyć okna poza widocznym obszarem na mniejszym.
            screen = QGuiApplication.primaryScreen()
            if screen:
                avail = screen.availableGeometry()
                w = min(w, avail.width())
                h = min(h, avail.height())
            self.resize(max(w, MIN_WINDOW_SIZE[0]), max(h, MIN_WINDOW_SIZE[1]))

        split = data.get("splitter")
        if isinstance(split, list) and len(split) == 2 \
                and all(isinstance(v, int) and v >= 0 for v in split) and sum(split) > 0:
            self.splitter.setSizes(split)

        if isinstance(data.get("wheel_inverted"), bool):
            self.preview.set_wheel_inverted(data["wheel_inverted"])
        if data.get("always_on_top") is True:
            # setChecked odpala toggled -> _toggle_always_on_top, wiec flaga
            # okna i wyglad przycisku ustawiaja sie same.
            self.btn_on_top.setChecked(True)
        self._page_number_cfg = self._sanitize_page_numbers(data.get("page_numbers"))
        last_dir = data.get("last_dir")
        if isinstance(last_dir, str) and os.path.isdir(last_dir):
            self._last_dir = last_dir

        if data.get("maximized") is True:
            self.showMaximized()

    def closeEvent(self, event):
        # Przy zmaksymalizowanym oknie zapisujemy rozmiar SPRZED maksymalizacji,
        # żeby po przywróceniu nie startowało na cały ekran przez przypadek.
        geom = self.normalGeometry() if self.isMaximized() else self.geometry()
        state = {
            "window": [geom.width(), geom.height()],
            "splitter": list(self.splitter.sizes()),
            "maximized": self.isMaximized(),
            "wheel_inverted": self.preview.wheel_inverted(),
            "always_on_top": self._always_on_top,
            "page_numbers": self._page_number_cfg,
            "last_dir": self._last_dir,
        }
        try:
            with open(self._window_state_file(), "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
        except Exception:
            pass  # brak zapisu ustawień nie może blokować zamknięcia programu
        super().closeEvent(event)

    def _history_file(self):
        return os.path.join(get_app_dir(), PROJECT_HISTORY_FILENAME)

    def _load_history(self, verify=True):
        """Historia projektów, domyślnie oczyszczona z nieistniejących plików.

        Weryfikacja dzieje się przy starcie i przy każdym rozwinięciu listy, a
        gdy coś odpadnie, plik historii jest od razu przepisywany — inaczej
        martwe wpisy (skasowane albo przeniesione projekty) zostawałyby na
        liście w nieskończoność. Przy okazji lecą duplikaty.
        """
        try:
            with open(self._history_file(), "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return []
        if not isinstance(data, list):
            return []

        entries = [p for p in data if isinstance(p, str)]
        if not verify:
            return entries

        kept, seen = [], set()
        for path in entries:
            key = os.path.normcase(os.path.abspath(path))
            if key in seen:
                continue
            seen.add(key)
            if os.path.exists(path):
                kept.append(path)
        if kept != entries:
            self._save_history(kept)
        return kept

    def _save_history(self, history):
        try:
            with open(self._history_file(), "w", encoding="utf-8") as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass  # brak zapisu historii nie może wywrócić głównej funkcji

    def _add_to_history(self, path):
        history = [p for p in self._load_history()
                   if os.path.normcase(p) != os.path.normcase(path)]
        history.insert(0, path)
        self._save_history(history[:PROJECT_HISTORY_LIMIT])
        self._refresh_project_history()

    def _remove_from_history(self, path):
        history = [p for p in self._load_history()
                   if os.path.normcase(p) != os.path.normcase(path)]
        self._save_history(history)
        self._refresh_project_history()

    def _projects_in_app_dir(self):
        """Projekty lezace obok programu (exe albo .py — patrz get_app_dir).

        Skanujemy tylko ten jeden katalog i tylko pliki .json, ktore naprawde
        wygladaja jak projekt (klucz "targets"), zeby na liscie nie ladowaly
        historia, stan okna ani obce pliki JSON.
        """
        app_dir = get_app_dir()
        skip = {os.path.normcase(PROJECT_HISTORY_FILENAME),
                os.path.normcase(WINDOW_STATE_FILENAME)}
        found = []
        try:
            names = sorted(os.listdir(app_dir))
        except OSError:
            return found
        for name in names:
            if not name.lower().endswith(".json") or os.path.normcase(name) in skip:
                continue
            path = os.path.join(app_dir, name)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                continue  # uszkodzony albo nie nasz plik — pomijamy po cichu
            if isinstance(data, dict) and isinstance(data.get("targets"), list):
                found.append(path)
        return found

    def _project_items(self):
        """Lista rozwijana: historia + projekty z katalogu programu, bez powtorzen."""
        items = list(self._load_history())
        seen = {os.path.normcase(os.path.abspath(p)) for p in items}
        for path in self._projects_in_app_dir():
            key = os.path.normcase(os.path.abspath(path))
            if key not in seen:
                seen.add(key)
                items.append(path)
        return items

    def _refresh_project_history(self):
        """Przebudowa listy projektow. Pole pokazuje TO, co faktycznie jest
        wczytane: bez projektu zostaje pozycja "(Niezapisany)" na indeksie 0.

        Wczesniej po clear()+addItems() combo samo wskakiwalo na pierwszy wpis
        historii, wiec przy starcie w polu widniala ostatnia sciezka, a drzewo
        pokazywalo pusty projekt — dopiero przelaczenie tam i z powrotem
        naprawialo niezgodnosc.
        """
        items = self._project_items()
        self.combo_project.blockSignals(True)
        self.combo_project.clear()
        self.combo_project.addItems([self.tr_("project_unsaved")] + items)
        self._select_project_item(self.project_file_path)
        self.combo_project.blockSignals(False)

    def _select_project_item(self, path):
        """Ustawia pole na dany projekt (albo na "(Niezapisany)" dla pustego).

        Wybor idzie po INDEKSIE, bo setCurrentText() Fluenta dziala tylko dla
        tekstu bedacego pozycja listy.
        """
        if not path:
            self.combo_project.setCurrentIndex(0)
            return
        key = os.path.normcase(os.path.abspath(path))
        for i in range(1, self.combo_project.count()):
            if os.path.normcase(os.path.abspath(self.combo_project.itemText(i))) == key:
                self.combo_project.setCurrentIndex(i)
                return
        self.combo_project.addItem(path)
        self.combo_project.setCurrentIndex(self.combo_project.count() - 1)

    def _on_history_selected(self, text):
        if text and os.path.isabs(text) and text != self.project_file_path:
            self._load_project_from_path(text)

    # --- KOMUNIKATY ---
    def _warn(self, title, body):
        InfoDialog(self, title, body, self.tr_("btn_ok")).exec()

    def _error(self, title, body):
        InfoDialog(self, title, body, self.tr_("btn_ok")).exec()


def system_accent_color():
    """Kolor wyrozniajacy z ustawien Windows (Personalizacja > Kolory).

    Zrodlem jest wartosc AccentColor w kluczu DWM rejestru — DWORD w kolejnosci ABGR, a nie
    ARGB, wiec skladowe R i B trzeba zamienic. Gdy klucza nie ma (inny system,
    polityka), bierzemy akcent z palety Qt, a na koncu zostaje kolor wlasny
    biblioteki. Zwraca None tylko wtedy, gdy nic nie da sie odczytac.
    """
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\DWM") as key:
            value = winreg.QueryValueEx(key, "AccentColor")[0]
        return QColor(value & 0xFF, (value >> 8) & 0xFF, (value >> 16) & 0xFF)
    except Exception:
        pass
    app = QApplication.instance()
    if app is not None:
        accent = app.palette().accent().color()
        if accent.isValid():
            return accent
    return None


def apply_system_accent():
    """Ustawia kolor motywu na akcent Windows. Bez tego qfluentwidgets
    maluje przelaczniki i przycisk scalania swoim morskim #009faa."""
    accent = system_accent_color()
    if accent is not None:
        setThemeColor(accent, save=False)


def main():
    QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    icon_path = os.path.join(get_resource_dir(), "app.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    setTheme(Theme.AUTO)
    apply_system_accent()
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
