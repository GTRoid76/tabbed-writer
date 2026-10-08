"""Tabbed Writer: a simple word processor where each tab is a chapter.

The book settings follow Amazon KDP's published guidance:
- eBooks: chapter titles use Heading 1 and are centered, body paragraphs
  have a 0.2 inch first-line indent with no extra space before or after,
  and Word (.docx) or EPUB files are uploaded.
- Paperbacks: standard trim sizes (6 x 9 in is the most common), mirrored
  inside/outside margins, and a print PDF.

Needs: Python 3, PySide6, and python-docx (for Word files).
LibreOffice is optional: it's only used to open .odt files.
Works on Linux and Windows.
"""

import html as html_text
import io
import json
import shutil
import subprocess
import sys
import tempfile
import uuid
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from PySide6.QtCore import QMarginsF, QSettings, QSizeF, QStandardPaths, Qt, QTimer
from PySide6.QtGui import (
    QAction,
    QBrush,
    QColor,
    QFont,
    QIcon,
    QKeySequence,
    QPageLayout,
    QPageSize,
    QPdfWriter,
    QTextBlockFormat,
    QTextCharFormat,
    QTextCursor,
    QTextDocument,
    QTextDocumentFragment,
    QTextDocumentWriter,
    QTextFormat,
    QTextFrameFormat,
    QTextLength,
    QTextTableFormat,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHeaderView,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMainWindow,
    QMenu,
    QMenuBar,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

DOCUMENT_FILE = None
SUGGESTED_NAME = "Untitled.twp"

# Where your writing is kept. Open and Save as start here.
DOCUMENTS = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
WRITING_FOLDER = Path(DOCUMENTS or Path.home() / "Documents") / "Tabbed Writer"
WRITING_FOLDER.mkdir(parents=True, exist_ok=True)
CONTENT_ROLE = Qt.ItemDataRole.UserRole
WORDS_ROLE = Qt.ItemDataRole.UserRole + 1

# "standard" documents behave like any word processor.
# "ebook" documents follow the KDP / B&N Press formatting rules.
document_kind = "standard"

app_settings = QSettings("TabbedWriter", "TabbedWriter")
DEFAULT_WORDS_PER_PAGE = 250

# The same three kinds of font Kindle Create offers (serif, sans serif,
# monospace). Linux uses the free Liberation fonts; Windows and Mac use
# Times New Roman, Arial and Courier New, which they were designed to match.
ON_LINUX = sys.platform.startswith("linux")
LINUX_FONTS = {"Serif": "Liberation Serif", "Sans serif": "Liberation Sans", "Monospace": "Liberation Mono"}
OTHER_FONTS = {"Serif": "Times New Roman", "Sans serif": "Arial", "Monospace": "Courier New"}
FONT_CHOICES = LINUX_FONTS if ON_LINUX else OTHER_FONTS
FONT_LABELS = {family: label for fonts in (LINUX_FONTS, OTHER_FONTS) for label, family in fonts.items()}
DEFAULT_FONT = FONT_CHOICES["Serif"]
DEFAULT_SIZE = 12
SIZE_CHOICES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"]

# KDP: first-line indent of 0.2 inch (Qt measures this at 96 dots per inch).
BOOK_INDENT = 0.2 * 96

# KDP paperback trim sizes (width, height in inches). 6 x 9 is the most common.
TRIM_SIZES = {
    "5 x 8 in": (5.0, 8.0),
    "5.25 x 8 in": (5.25, 8.0),
    "5.5 x 8.5 in": (5.5, 8.5),
    "6 x 9 in": (6.0, 9.0),
    "7 x 10 in": (7.0, 10.0),
    "8.5 x 11 in (Letter)": (8.5, 11.0),
    "8.27 x 11.69 in (A4)": (8.27, 11.69),
}
# Normal documents: Letter paper, 1 inch margins, like most word processors.
STANDARD_PAGE = {"trim": "8.5 x 11 in (Letter)", "top": 1.0, "bottom": 1.0, "inside": 1.0, "outside": 1.0}
# eBooks: KDP's most common paperback size, in case a print copy is wanted later.
DEFAULT_PAGE = {"trim": "6 x 9 in", "top": 0.75, "bottom": 0.75, "inside": 0.75, "outside": 0.5}
page_settings = dict(STANDARD_PAGE)
SCREEN_DPI = 96

# What "Save as" can make. Everything except .twp is a copy.
SAVE_TYPES = {
    "Tabbed Writer document (*.twp)": ".twp",
    "Word Document (*.docx)": ".docx",
    "EPUB (*.epub)": ".epub",
    "PDF (*.pdf)": ".pdf",
    "OpenDocument text (*.odt)": ".odt",
    "Plain text (*.txt)": ".txt",
}


# ---------------------------------------------------------------------------
# Tabs and saving
# ---------------------------------------------------------------------------

def make_item(parent, item_data):
    if parent is None:
        item = QTreeWidgetItem(sections, [item_data["title"]])
    else:
        item = QTreeWidgetItem(parent, [item_data["title"]])

    item.setData(0, CONTENT_ROLE, item_data.get("text", ""))
    item.setData(0, WORDS_ROLE, words_in_html(item_data.get("text", "")))

    for child_data in item_data.get("children", []):
        make_item(item, child_data)

    return item


def item_to_data(item):
    return {
        "title": item.text(0),
        "text": item.data(0, CONTENT_ROLE) or "",
        "children": [
            item_to_data(item.child(index))
            for index in range(item.childCount())
        ],
    }


saved_snapshot = ""


def count_words(text):
    return len(text.split())


def words_in_html(html):
    if not html:
        return 0
    holder = QTextDocument()
    holder.setHtml(html)
    return count_words(holder.toPlainText())


def words_per_page():
    try:
        return max(50, int(app_settings.value("words_per_page", DEFAULT_WORDS_PER_PAGE)))
    except (TypeError, ValueError):
        return DEFAULT_WORDS_PER_PAGE


def page_position():
    # Lay the whole document out on pages (like printing it) to find
    # how many pages there are and which page the cursor is on.
    text_width, text_height = page_text_size()
    layout_document = QTextDocument()
    layout_document.setDefaultFont(QFont(DEFAULT_FONT, DEFAULT_SIZE))
    layout_document.setDocumentMargin(0)
    cursor = QTextCursor(layout_document)
    current = sections.currentItem()
    cursor_position = 0
    for index in range(sections.topLevelItemCount()):
        item = sections.topLevelItem(index)
        if index > 0:
            cursor.insertBlock(QTextBlockFormat(), QTextCharFormat())
        if item is current:
            start = cursor.position()
            cursor.insertFragment(QTextDocumentFragment(editor.document()))
            cursor_position = start + editor.textCursor().position()
        else:
            html = item.data(0, CONTENT_ROLE)
            if html:
                cursor.insertFragment(QTextDocumentFragment.fromHtml(html))
    layout_document.setPageSize(QSizeF(text_width, text_height))
    total = max(1, layout_document.pageCount())

    cursor_position = min(cursor_position, layout_document.characterCount() - 1)
    block = layout_document.findBlock(cursor_position)
    y = layout_document.documentLayout().blockBoundingRect(block).y()
    line = block.layout().lineForTextPosition(cursor_position - block.position())
    if line.isValid():
        y += line.y()
    page = min(total, int(y // text_height) + 1)
    return page, total


def words_by_tab():
    current = sections.currentItem()
    counts = []
    for index in range(sections.topLevelItemCount()):
        item = sections.topLevelItem(index)
        if item is current:
            words = count_words(editor.toPlainText())
        else:
            words = item.data(0, WORDS_ROLE) or 0
        counts.append((item.text(0), words))
    return counts


def estimated_pages(total):
    return -(-total // words_per_page()) if total else 0


def plural(number, word):
    return f"{number:,} {word}" + ("" if number == 1 else "s")


def fill_word_table(table, counts):
    table.setRowCount(len(counts))
    for row, (title, words) in enumerate(counts):
        table.setItem(row, 0, QTableWidgetItem(title))
        number = QTableWidgetItem(f"{words:,}")
        number.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        table.setItem(row, 1, number)


def make_word_table():
    table = QTableWidget(0, 2)
    table.setHorizontalHeaderLabels(["Tab", "Words"])
    table.verticalHeader().setVisible(False)
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
    table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
    table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
    return table


def update_counts():
    counts = words_by_tab()
    total = sum(words for _, words in counts)
    if not word_panel.isHidden():
        fill_word_table(panel_table, counts)
        panel_total.setText(f"Total: {plural(total, 'word')}")
    if document_kind == "ebook":
        # eBooks have no fixed pages, so this is an estimate from words per page.
        count_label.setText(
            f"eBook  ·  {plural(total, 'word')}  ·  about {plural(estimated_pages(total), 'page')}"
        )
    else:
        page, pages = page_position()
        count_label.setText(f"Page {page:,} of {pages:,}  ·  {plural(total, 'word')}")


def show_word_count():
    counts = words_by_tab()
    total = sum(words for _, words in counts)

    dialog = QDialog(window)
    dialog.setWindowTitle("Word count")
    column = QVBoxLayout(dialog)

    table = make_word_table()
    fill_word_table(table, counts)
    column.addWidget(table)

    summary = f"Total: {plural(total, 'word')}"
    if document_kind == "ebook":
        summary += (
            f"\nAbout {plural(estimated_pages(total), 'page')} "
            f"(at {words_per_page():,} words per page)"
        )
    else:
        summary += f"\n{plural(page_position()[1], 'page')}"
    summary_label = QLabel(summary)
    summary_font = summary_label.font()
    summary_font.setBold(True)
    summary_label.setFont(summary_font)
    column.addWidget(summary_label)

    keep_showing = QCheckBox("Show in the tab panel while typing")
    keep_showing.setChecked(show_word_count_action.isChecked())
    keep_showing.toggled.connect(show_word_count_action.setChecked)
    column.addWidget(keep_showing)

    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
    buttons.rejected.connect(dialog.reject)
    column.addWidget(buttons)

    dialog.resize(380, 360)
    dialog.exec()


def update_counts_soon():
    # Wait for a short pause in typing before recounting.
    count_timer.start()


def change_words_per_page():
    number, accepted = QInputDialog.getInt(
        window,
        "eBook words per page",
        "Words per page, for the eBook page estimate.\n"
        "KDP and B&N don't publish a number; 250 is the common rule of thumb.",
        words_per_page(),
        100,
        600,
        10,
    )
    if accepted:
        app_settings.setValue("words_per_page", number)
        update_counts()


def editor_html():
    if editor.toPlainText() == "" and editor.document().blockCount() == 1:
        return ""
    return editor.toHtml()


def current_snapshot():
    current = sections.currentItem()
    if current is not None:
        current.setData(0, CONTENT_ROLE, editor_html())
    data = {
        "format": "Tabbed Writer",
        "version": 1,
        "kind": document_kind,
        "page": page_settings,
        "tabs": [
            item_to_data(sections.topLevelItem(index))
            for index in range(sections.topLevelItemCount())
        ],
    }
    return json.dumps(data, indent=2)


def save_document():
    global saved_snapshot
    if DOCUMENT_FILE is None:
        save_as_document()
        return
    snapshot = current_snapshot()
    DOCUMENT_FILE.write_text(snapshot, encoding="utf-8")
    saved_snapshot = snapshot
    window.statusBar().showMessage("Saved", 3000)


def ok_to_discard():
    if current_snapshot() == saved_snapshot:
        return True
    answer = QMessageBox.question(
        window,
        "Unsaved changes",
        "Save changes to this document first?",
        QMessageBox.StandardButton.Save
        | QMessageBox.StandardButton.Discard
        | QMessageBox.StandardButton.Cancel,
    )
    if answer == QMessageBox.StandardButton.Save:
        save_document()
        return current_snapshot() == saved_snapshot
    return answer == QMessageBox.StandardButton.Discard


def page_text_size():
    width, height = TRIM_SIZES.get(page_settings["trim"], TRIM_SIZES[STANDARD_PAGE["trim"]])
    text_width = (width - page_settings["inside"] - page_settings["outside"]) * SCREEN_DPI
    text_height = (height - page_settings["top"] - page_settings["bottom"]) * SCREEN_DPI
    return max(text_width, 100), max(text_height, 100)


def apply_page_look():
    # The writing area looks like a sheet of paper, using the Page setup size and margins.
    width, _ = TRIM_SIZES.get(page_settings["trim"], TRIM_SIZES[STANDARD_PAGE["trim"]])
    root = editor.document().rootFrame()
    frame_format = root.frameFormat()
    frame_format.setLeftMargin(page_settings["inside"] * SCREEN_DPI)
    frame_format.setRightMargin(page_settings["outside"] * SCREEN_DPI)
    frame_format.setTopMargin(page_settings["top"] * SCREEN_DPI)
    frame_format.setBottomMargin(page_settings["bottom"] * SCREEN_DPI)
    root.setFrameFormat(frame_format)
    editor.setMaximumWidth(int(width * SCREEN_DPI) + editor.verticalScrollBar().sizeHint().width())


def use_book_paragraph_if_empty():
    # In eBooks, new empty tabs start with the KDP paragraph style (0.2 inch indent).
    if document_kind != "ebook":
        return
    if editor.toPlainText() == "" and editor.document().blockCount() == 1:
        book_paragraph = QTextBlockFormat()
        book_paragraph.setTextIndent(BOOK_INDENT)
        book_paragraph.setTopMargin(0)
        book_paragraph.setBottomMargin(0)
        editor.textCursor().setBlockFormat(book_paragraph)


def switch_section(current, previous):
    if previous is not None:
        previous.setData(0, CONTENT_ROLE, editor_html())
        previous.setData(0, WORDS_ROLE, count_words(editor.toPlainText()))

    if current is not None:
        editor.setHtml(current.data(0, CONTENT_ROLE) or "")
        apply_page_look()
        use_book_paragraph_if_empty()
    update_counts()


def add_main_tab():
    if document_kind == "ebook":
        label, suggested = "Chapter name:", "Chapter"
    else:
        label, suggested = "Tab name:", "Tab"
    title, accepted = QInputDialog.getText(
        window, "Add tab", label, text=f"{suggested} {sections.topLevelItemCount() + 1}"
    )
    if not accepted or not title.strip():
        return
    item = QTreeWidgetItem(sections, [title.strip()])
    item.setData(0, CONTENT_ROLE, "")
    item.setData(0, WORDS_ROLE, 0)
    sections.setCurrentItem(item)


def rename_section():
    item = sections.currentItem()
    if item is None:
        return

    title, accepted = QInputDialog.getText(
        window,
        "Rename tab",
        "Tab name:",
        text=item.text(0),
    )
    if accepted and title.strip():
        item.setText(0, title.strip())
        update_counts_soon()


def remove_section():
    item = sections.currentItem()
    if item is None:
        return

    if sections.topLevelItemCount() <= 1:
        QMessageBox.information(window, "Delete tab", "A document needs at least one tab.")
        return

    answer = QMessageBox.question(
        window,
        "Delete tab",
        f"Delete “{item.text(0)}” and everything written in it?",
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
    )
    if answer != QMessageBox.StandardButton.Yes:
        return

    index = sections.indexOfTopLevelItem(item)

    sections.blockSignals(True)
    sections.takeTopLevelItem(index)
    count = sections.topLevelItemCount()
    next_item = sections.topLevelItem(min(index, count - 1))
    sections.setCurrentItem(next_item)
    sections.blockSignals(False)

    editor.setHtml(next_item.data(0, CONTENT_ROLE) or "")
    apply_page_look()
    use_book_paragraph_if_empty()
    update_counts()


def fill_tabs(tabs):
    sections.blockSignals(True)
    sections.clear()
    for tab in tabs:
        make_item(None, tab)
    if sections.topLevelItemCount() == 0:
        make_item(None, {"title": "Tab 1", "text": ""})
    sections.blockSignals(False)
    editor.clear()
    sections.setCurrentItem(sections.topLevelItem(0))


def read_page_settings(saved_page):
    settings = dict(STANDARD_PAGE)
    for key in STANDARD_PAGE:
        if key in saved_page:
            settings[key] = saved_page[key]
    if settings["trim"] not in TRIM_SIZES:
        settings["trim"] = STANDARD_PAGE["trim"]
    return settings


def set_document_kind(kind):
    global document_kind
    document_kind = kind
    ebook_action.blockSignals(True)
    ebook_action.setChecked(kind == "ebook")
    ebook_action.blockSignals(False)


def load_document(path):
    global page_settings
    saved = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(saved, list):
        tabs = saved
        page_settings = dict(STANDARD_PAGE)
        set_document_kind("standard")
    else:
        tabs = saved.get("tabs", [])
        page_settings = read_page_settings(saved.get("page", {}))
        set_document_kind("ebook" if saved.get("kind") == "ebook" else "standard")
    apply_page_look()
    fill_tabs(tabs)


# ---------------------------------------------------------------------------
# LibreOffice (optional: only used to open .odt files)
# ---------------------------------------------------------------------------

def find_libreoffice():
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    for place in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if Path(place).exists():
            return place
    return None


def libreoffice_convert(source, target_format, out_dir):
    office = find_libreoffice()
    if office is None:
        raise RuntimeError(
            "Opening .odt files uses LibreOffice, which isn't installed on this computer.\n\n"
            "You can get it free from libreoffice.org, or save the file as a Word "
            "document (.docx) in the program that made it, then open that."
        )
    QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
    try:
        with tempfile.TemporaryDirectory() as profile:
            subprocess.run(
                [
                    office,
                    "--headless",
                    f"-env:UserInstallation={Path(profile).as_uri()}",
                    "--convert-to",
                    target_format,
                    "--outdir",
                    str(out_dir),
                    str(source),
                ],
                capture_output=True,
                timeout=180,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
    finally:
        QApplication.restoreOverrideCursor()


# ---------------------------------------------------------------------------
# Book settings inside .odt files
# ---------------------------------------------------------------------------

ODF_NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "dc": "http://purl.org/dc/elements/1.1/",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
}


def odf_tag(prefix, name):
    return f"{{{ODF_NS[prefix]}}}{name}"


def read_xml(data):
    for prefix, uri in ODF_NS.items():
        ET.register_namespace(prefix, uri)
    for _, (prefix, uri) in ET.iterparse(io.BytesIO(data), events=["start-ns"]):
        if prefix:
            ET.register_namespace(prefix, uri)
    return ET.fromstring(data)


def xml_bytes(root):
    return b'<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(root)


def odf_child(parent, prefix, name, before=None):
    found = parent.find(odf_tag(prefix, name))
    if found is None:
        found = ET.Element(odf_tag(prefix, name))
        later = [parent.find(odf_tag(p, n)) for p, n in (before or [])]
        later = [element for element in later if element is not None]
        if later:
            parent.insert(list(parent).index(later[0]), found)
        else:
            parent.append(found)
    return found


def fix_odt_styles(root, page):
    styles = odf_child(root, "office", "styles",
                       before=[("office", "automatic-styles"), ("office", "master-styles")])
    automatic = odf_child(root, "office", "automatic-styles", before=[("office", "master-styles")])
    masters = odf_child(root, "office", "master-styles")

    names = {element.get(odf_tag("style", "name")) for element in styles}
    if "Heading_20_1" not in names:
        heading = ET.SubElement(styles, odf_tag("style", "style"), {
            odf_tag("style", "name"): "Heading_20_1",
            odf_tag("style", "display-name"): "Heading 1",
            odf_tag("style", "family"): "paragraph",
            odf_tag("style", "default-outline-level"): "1",
            odf_tag("style", "class"): "text",
        })
        ET.SubElement(heading, odf_tag("style", "paragraph-properties"), {
            odf_tag("fo", "text-align"): "center",
            odf_tag("fo", "margin-top"): "0in",
            odf_tag("fo", "margin-bottom"): "0.25in",
            odf_tag("fo", "keep-with-next"): "always",
        })
        ET.SubElement(heading, odf_tag("style", "text-properties"), {
            odf_tag("fo", "font-size"): "18pt",
            odf_tag("fo", "font-weight"): "bold",
        })

    for old in automatic.findall(odf_tag("style", "page-layout")):
        if old.get(odf_tag("style", "name")) == "TWPageLayout":
            automatic.remove(old)
    width, height = TRIM_SIZES.get(page["trim"], TRIM_SIZES[DEFAULT_PAGE["trim"]])
    layout = ET.SubElement(automatic, odf_tag("style", "page-layout"), {
        odf_tag("style", "name"): "TWPageLayout",
        odf_tag("style", "page-usage"): "mirrored",
    })
    ET.SubElement(layout, odf_tag("style", "page-layout-properties"), {
        odf_tag("fo", "page-width"): f"{width}in",
        odf_tag("fo", "page-height"): f"{height}in",
        odf_tag("style", "print-orientation"): "portrait",
        odf_tag("fo", "margin-top"): f"{page['top']}in",
        odf_tag("fo", "margin-bottom"): f"{page['bottom']}in",
        odf_tag("fo", "margin-left"): f"{page['inside']}in",
        odf_tag("fo", "margin-right"): f"{page['outside']}in",
    })

    master_pages = masters.findall(odf_tag("style", "master-page"))
    if not master_pages:
        master_pages = [ET.SubElement(masters, odf_tag("style", "master-page"),
                                      {odf_tag("style", "name"): "Standard"})]
    for master in master_pages:
        master.set(odf_tag("style", "page-layout-name"), "TWPageLayout")


def fix_odt_headings(root):
    automatic = root.find(odf_tag("office", "automatic-styles"))
    by_name = {}
    if automatic is not None:
        by_name = {element.get(odf_tag("style", "name")): element for element in automatic}
    for heading in root.iter(odf_tag("text", "h")):
        if heading.get(odf_tag("text", "outline-level")) != "1":
            continue
        style = by_name.get(heading.get(odf_tag("text", "style-name")))
        if style is None:
            heading.set(odf_tag("text", "style-name"), "Heading_20_1")
        elif style.get(odf_tag("style", "parent-style-name")) is None:
            style.set(odf_tag("style", "parent-style-name"), "Heading_20_1")


def finish_odt(file_name, page, title):
    """Add the page setup, a real Heading 1 style, and the book title to an .odt file."""
    with zipfile.ZipFile(file_name) as source:
        parts = [(info.filename, source.read(info.filename)) for info in source.infolist()]
    contents = dict(parts)
    added = []

    if "styles.xml" in contents:
        styles_root = read_xml(contents["styles.xml"])
    else:
        read_xml(b"<x/>")
        styles_root = ET.Element(odf_tag("office", "document-styles"), {odf_tag("office", "version"): "1.2"})
        added.append("styles.xml")
    fix_odt_styles(styles_root, page)
    contents["styles.xml"] = xml_bytes(styles_root)

    content_root = read_xml(contents["content.xml"])
    fix_odt_headings(content_root)
    contents["content.xml"] = xml_bytes(content_root)

    if "meta.xml" not in contents:
        meta_root = ET.Element(odf_tag("office", "document-meta"), {odf_tag("office", "version"): "1.2"})
        meta = ET.SubElement(meta_root, odf_tag("office", "meta"))
        ET.SubElement(meta, odf_tag("dc", "title")).text = title
        contents["meta.xml"] = xml_bytes(meta_root)
        added.append("meta.xml")

    if added:
        entries = "".join(
            f' <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="{name}"/>\n'
            for name in added
        ).encode()
        contents["META-INF/manifest.xml"] = contents["META-INF/manifest.xml"].replace(
            b"</manifest:manifest>", entries + b"</manifest:manifest>"
        )

    order = [name for name, _ in parts] + added
    with zipfile.ZipFile(file_name, "w") as target:
        for name in order:
            if name == "mimetype":
                target.writestr(name, contents[name], compress_type=zipfile.ZIP_STORED)
            else:
                target.writestr(name, contents[name], compress_type=zipfile.ZIP_DEFLATED)


# ---------------------------------------------------------------------------
# Word and EPUB files, written directly (no LibreOffice needed)
#
# A "book" here is a plain list of items:
#   {"type": "paragraph", "heading": 0 or 1, "align": "left"/"center"/"right"/"justify",
#    "indent": True/False, "page_break": True/False,
#    "runs": [{"text", "bold", "italic", "underline", "color", "highlight", "font", "size"}]}
#   {"type": "table", "border": "#rrggbb" or None,
#    "rows": [[{"background": "#rrggbb" or None, "paragraphs": [paragraph, ...]}, ...], ...]}
# ---------------------------------------------------------------------------

WORD_FONTS = {
    "Liberation Serif": "Times New Roman",
    "Liberation Sans": "Arial",
    "Liberation Mono": "Courier New",
}


GENERIC_FONTS = {
    "Liberation Serif": "serif",
    "Times New Roman": "serif",
    "Liberation Sans": "sans-serif",
    "Arial": "sans-serif",
    "Liberation Mono": "monospace",
    "Courier New": "monospace",
}


def word_font(family):
    return WORD_FONTS.get(family, family)


# ----- Word (.docx) --------------------------------------------------------

def write_docx(book, file_name, page, trim_sizes, title, ebook):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor

    alignments = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }

    document = Document()
    document.core_properties.title = title

    width, height = trim_sizes[page["trim"]]
    section = document.sections[0]
    section.page_width = Inches(width)
    section.page_height = Inches(height)
    section.top_margin = Inches(page["top"])
    section.bottom_margin = Inches(page["bottom"])
    section.left_margin = Inches(page["inside"])
    section.right_margin = Inches(page["outside"])
    if ebook:
        mirror = OxmlElement("w:mirrorMargins")
        document.settings.element.append(mirror)

    normal = document.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0 if ebook else 6)

    heading_style = document.styles["Heading 1"]
    heading_style.font.size = Pt(18)
    heading_style.font.bold = True
    heading_style.font.color.rgb = RGBColor(0, 0, 0)
    heading_style.font.name = "Times New Roman"
    heading_style.paragraph_format.space_before = Pt(0)
    heading_style.paragraph_format.space_after = Pt(18)

    def shade(properties, color):
        shading = OxmlElement("w:shd")
        shading.set(qn("w:val"), "clear")
        shading.set(qn("w:color"), "auto")
        shading.set(qn("w:fill"), color.lstrip("#").upper())
        properties.append(shading)

    def fill_paragraph(paragraph, item):
        if item["heading"]:
            paragraph.style = document.styles["Heading 1"]
        paragraph.alignment = alignments.get(item["align"], WD_ALIGN_PARAGRAPH.LEFT)
        if item["indent"]:
            paragraph.paragraph_format.first_line_indent = Inches(0.2)
        if item["page_break"]:
            paragraph.paragraph_format.page_break_before = True
        for run_data in item["runs"]:
            pieces = run_data["text"].split(" ")
            for number, piece in enumerate(pieces):
                if number:
                    paragraph.add_run().add_break(WD_BREAK.LINE)
                if not piece:
                    continue
                run = paragraph.add_run(piece)
                run.bold = run_data["bold"] or None
                run.italic = run_data["italic"] or None
                run.underline = run_data["underline"] or None
                if run_data["font"]:
                    run.font.name = word_font(run_data["font"])
                if run_data["size"]:
                    run.font.size = Pt(run_data["size"])
                if run_data["color"]:
                    run.font.color.rgb = RGBColor.from_string(run_data["color"].lstrip("#").upper())
                if run_data["highlight"]:
                    shade(run._element.get_or_add_rPr(), run_data["highlight"])

    def set_table_borders(table, color):
        table_properties = table._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            line = OxmlElement(f"w:{side}")
            line.set(qn("w:val"), "single")
            line.set(qn("w:sz"), "4")
            line.set(qn("w:space"), "0")
            line.set(qn("w:color"), (color or "#000000").lstrip("#").upper())
            borders.append(line)
        table_properties.append(borders)

    for item in book:
        if item["type"] == "paragraph":
            fill_paragraph(document.add_paragraph(), item)
            continue
        rows = item["rows"]
        columns = max((len(row) for row in rows), default=0)
        if not columns:
            continue
        table = document.add_table(rows=len(rows), cols=columns)
        set_table_borders(table, item["border"])
        for row_number, row in enumerate(rows):
            for column_number, cell_data in enumerate(row):
                cell = table.cell(row_number, column_number)
                if cell_data["background"]:
                    shade(cell._tc.get_or_add_tcPr(), cell_data["background"])
                paragraphs = cell_data["paragraphs"] or []
                for number, paragraph_data in enumerate(paragraphs):
                    paragraph = cell.paragraphs[0] if number == 0 else cell.add_paragraph()
                    fill_paragraph(paragraph, dict(paragraph_data, indent=False, page_break=False))
        document.add_paragraph()

    # Word documents always start with one empty paragraph; remove it.
    first = document.paragraphs[0]
    if not first.text and len(document.paragraphs) > 1:
        first._element.getparent().remove(first._element)

    document.save(file_name)


# ----- EPUB ----------------------------------------------------------------

EPUB_CSS = """body { font-family: serif; margin: 0 5%; }
p { margin: 0 0 0.6em 0; }
p.book { margin: 0; text-indent: 1.2em; }
p.noindent { text-indent: 0; }
h1 { font-size: 1.5em; font-weight: bold; margin: 1em 0 1em 0; }
h1.center { text-align: center; }
.center { text-align: center; }
.right { text-align: right; }
.justify { text-align: justify; }
.u { text-decoration: underline; }
table { border-collapse: collapse; width: 100%; margin: 0.6em 0; }
td { border: 1px solid #000; padding: 0.25em; vertical-align: top; }
td p { margin: 0; text-indent: 0; }
"""


def escape(text):
    return html_text.escape(text, quote=True)


def epub_runs(runs):
    out = []
    for run in runs:
        text = "<br/>".join(escape(piece) for piece in run["text"].split(" "))
        if not text:
            continue
        styles = []
        if run["color"]:
            styles.append(f"color: {run['color']}")
        if run["highlight"]:
            styles.append(f"background-color: {run['highlight']}")
        generic = GENERIC_FONTS.get(run["font"] or "", "serif")
        if generic != "serif":
            styles.append(f"font-family: {generic}")
        if run["underline"]:
            text = f'<span class="u">{text}</span>'
        if run["italic"]:
            text = f"<em>{text}</em>"
        if run["bold"]:
            text = f"<strong>{text}</strong>"
        if styles:
            text = f'<span style="{"; ".join(styles)}">{text}</span>'
        out.append(text)
    return "".join(out)


def epub_paragraph(item, ebook):
    content = epub_runs(item["runs"])
    classes = []
    if item["heading"]:
        if ebook or item["align"] == "center":
            classes.append("center")
        attribute = f' class="{" ".join(classes)}"' if classes else ""
        return f"<h1{attribute}>{content or '&#160;'}</h1>"
    if ebook:
        # KDP: no space between paragraphs; indented unless centered or right-aligned.
        classes.append("book")
        if not item["indent"]:
            classes.append("noindent")
    if item["align"] in ("center", "right", "justify"):
        classes.append(item["align"])
    attribute = f' class="{" ".join(classes)}"' if classes else ""
    return f"<p{attribute}>{content or '&#160;'}</p>"


def epub_table(item, ebook):
    border = item["border"] or "#000000"
    lines = ["<table>"]
    for row in item["rows"]:
        lines.append("<tr>")
        for cell in row:
            style = [f"border-color: {border}"]
            if cell["background"]:
                style.append(f"background-color: {cell['background']}")
            inside = "".join(
                epub_paragraph(dict(p, heading=0, indent=False), False) for p in cell["paragraphs"]
            ) or "<p>&#160;</p>"
            lines.append(f'<td style="{"; ".join(style)}">{inside}</td>')
        lines.append("</tr>")
    lines.append("</table>")
    return "\n".join(lines)


def split_chapters(book, fallback_title):
    chapters = []
    for item in book:
        starts_chapter = item["type"] == "paragraph" and item["heading"] == 1
        if starts_chapter or not chapters:
            title = "".join(run["text"] for run in item["runs"]).strip() if starts_chapter else ""
            chapters.append({"title": title or fallback_title, "items": []})
        chapters[-1]["items"].append(item)
    return chapters or [{"title": fallback_title, "items": []}]


def xhtml_page(title, body):
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        "<!DOCTYPE html>\n"
        '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" '
        'lang="en" xml:lang="en">\n'
        f"<head><meta charset=\"utf-8\"/><title>{escape(title)}</title>"
        '<link rel="stylesheet" type="text/css" href="style.css"/></head>\n'
        f"<body>\n{body}\n</body>\n</html>\n"
    )


def write_epub(book, file_name, title, author, ebook):
    chapters = split_chapters(book, title)
    book_id = f"urn:uuid:{uuid.uuid4()}"
    modified = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    pages = []
    for number, chapter in enumerate(chapters, start=1):
        body = "\n".join(
            epub_paragraph(item, ebook) if item["type"] == "paragraph" else epub_table(item, ebook)
            for item in chapter["items"]
        )
        pages.append((f"chapter{number:03}.xhtml", chapter["title"], xhtml_page(chapter["title"], body)))

    nav_items = "\n".join(
        f'<li><a href="{name}">{escape(chapter_title)}</a></li>' for name, chapter_title, _ in pages
    )
    nav = xhtml_page(
        "Contents",
        f'<nav epub:type="toc" id="toc"><h1>Contents</h1><ol>\n{nav_items}\n</ol></nav>',
    )

    manifest = "\n".join(
        f'<item id="c{number}" href="{name}" media-type="application/xhtml+xml"/>'
        for number, (name, _, _) in enumerate(pages, start=1)
    )
    spine = "\n".join(f'<itemref idref="c{number}"/>' for number in range(1, len(pages) + 1))
    creator = f"<dc:creator>{escape(author)}</dc:creator>\n" if author else ""
    opf = f"""<?xml version="1.0" encoding="utf-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="bookid" xml:lang="en">
<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:identifier id="bookid">{book_id}</dc:identifier>
<dc:title>{escape(title)}</dc:title>
{creator}<dc:language>en</dc:language>
<meta property="dcterms:modified">{modified}</meta>
</metadata>
<manifest>
<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>
<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>
<item id="css" href="style.css" media-type="text/css"/>
{manifest}
</manifest>
<spine toc="ncx">
{spine}
</spine>
</package>
"""

    nav_points = "\n".join(
        f'<navPoint id="n{number}" playOrder="{number}"><navLabel><text>{escape(chapter_title)}</text>'
        f'</navLabel><content src="{name}"/></navPoint>'
        for number, (name, chapter_title, _) in enumerate(pages, start=1)
    )
    ncx = f"""<?xml version="1.0" encoding="utf-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
<head><meta name="dtb:uid" content="{book_id}"/></head>
<docTitle><text>{escape(title)}</text></docTitle>
<navMap>
{nav_points}
</navMap>
</ncx>
"""

    container = """<?xml version="1.0" encoding="utf-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
<rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles>
</container>
"""

    with zipfile.ZipFile(file_name, "w") as target:
        target.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)
        target.writestr("META-INF/container.xml", container, compress_type=zipfile.ZIP_DEFLATED)
        target.writestr("OEBPS/content.opf", opf, compress_type=zipfile.ZIP_DEFLATED)
        target.writestr("OEBPS/nav.xhtml", nav, compress_type=zipfile.ZIP_DEFLATED)
        target.writestr("OEBPS/toc.ncx", ncx, compress_type=zipfile.ZIP_DEFLATED)
        target.writestr("OEBPS/style.css", EPUB_CSS, compress_type=zipfile.ZIP_DEFLATED)
        for name, _, page in pages:
            target.writestr(f"OEBPS/{name}", page, compress_type=zipfile.ZIP_DEFLATED)


# ----- Reading Word files --------------------------------------------------

def docx_to_html(file_name):
    """Turn a Word file into simple HTML. Heading 1 paragraphs become <h1>."""
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    document = Document(file_name)

    def paragraph_html(paragraph):
        style_name = (paragraph.style.name if paragraph.style is not None else "") or ""
        tag = "h1" if style_name in ("Heading 1", "Title") else "p"
        alignment = paragraph.alignment
        align = ""
        if alignment is not None:
            name = str(alignment).split(".")[-1].split(" ")[0].lower()
            if name.startswith("center"):
                align = "center"
            elif name.startswith("right"):
                align = "right"
            elif name.startswith("justify"):
                align = "justify"
        styles = []
        if align:
            styles.append(f"text-align: {align}")
        if paragraph.paragraph_format.page_break_before:
            styles.append("page-break-before: always")
        pieces = []
        for run in paragraph.runs:
            text = escape(run.text).replace("\n", "<br/>").replace("\t", "&#160;&#160;&#160;&#160;")
            if not text:
                continue
            run_styles = []
            if run.font.color is not None and run.font.color.type is not None and run.font.color.rgb is not None:
                run_styles.append(f"color: #{run.font.color.rgb}")
            if run.font.size:
                run_styles.append(f"font-size: {run.font.size.pt:g}pt")
            if run.underline:
                text = f"<u>{text}</u>"
            if run.italic:
                text = f"<i>{text}</i>"
            if run.bold:
                text = f"<b>{text}</b>"
            if run_styles:
                text = f'<span style="{"; ".join(run_styles)}">{text}</span>'
            pieces.append(text)
        style_attribute = f' style="{"; ".join(styles)}"' if styles else ""
        return f"<{tag}{style_attribute}>{''.join(pieces)}</{tag}>"

    def table_html(table):
        rows = []
        for row in table.rows:
            cells = "".join(
                "<td>" + "".join(paragraph_html(p) for p in cell.paragraphs) + "</td>"
                for cell in row.cells
            )
            rows.append(f"<tr>{cells}</tr>")
        return '<table border="1" cellspacing="0" cellpadding="4" width="100%">' + "".join(rows) + "</table>"

    parts = []
    body = document.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            parts.append(paragraph_html(Paragraph(child, document)))
        elif tag == "tbl":
            parts.append(table_html(Table(child, document)))
    return "<html><body>" + "\n".join(parts) + "</body></html>"


# ----- Turning the document into a "book" list ------------------------------

def brush_color(text_format, property_id):
    if text_format.hasProperty(property_id):
        brush = text_format.brushProperty(property_id)
        if brush.style() != Qt.BrushStyle.NoBrush:
            return brush.color().name()
    return None


def alignment_name(block_format):
    alignment = block_format.alignment()
    if alignment & Qt.AlignmentFlag.AlignHCenter:
        return "center"
    if alignment & Qt.AlignmentFlag.AlignRight:
        return "right"
    if alignment & Qt.AlignmentFlag.AlignJustify:
        return "justify"
    return "left"


def block_to_paragraph(block):
    block_format = block.blockFormat()
    encoded = block.text().encode("utf-16-le")
    runs = []
    for part in block.textFormats():
        piece = encoded[2 * part.start: 2 * (part.start + part.length)].decode("utf-16-le", errors="replace")
        if not piece:
            continue
        text_format = part.format
        font = text_format.font()
        family = font.family()
        runs.append({
            "text": piece,
            "bold": font.bold(),
            "italic": font.italic(),
            "underline": text_format.fontUnderline(),
            "color": brush_color(text_format, QTextFormat.Property.ForegroundBrush),
            "highlight": brush_color(text_format, QTextFormat.Property.BackgroundBrush),
            "font": family if family in FONT_LABELS and FONT_LABELS[family] != "Serif" else None,
            "size": text_format.fontPointSize() if text_format.fontPointSize() > 0 else 0,
        })
    return {
        "type": "paragraph",
        "heading": 1 if block_format.headingLevel() == 1 else 0,
        "align": alignment_name(block_format),
        "indent": block_format.textIndent() > 0,
        "page_break": bool(block_format.pageBreakPolicy() & QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore),
        "runs": runs,
    }


def table_to_item(document, table):
    rows = []
    for row in range(table.rows()):
        cells = []
        for column in range(table.columns()):
            cell = table.cellAt(row, column)
            paragraphs = []
            block = document.findBlock(cell.firstPosition())
            while block.isValid() and block.position() <= cell.lastPosition():
                paragraphs.append(block_to_paragraph(block))
                block = block.next()
            cells.append({
                "background": brush_color(cell.format(), QTextFormat.Property.BackgroundBrush),
                "paragraphs": paragraphs,
            })
        rows.append(cells)
    return {
        "type": "table",
        "border": brush_color(table.format(), QTextFormat.Property.FrameBorderBrush),
        "rows": rows,
    }


def document_to_book(document):
    book = []
    block = document.begin()
    while block.isValid():
        table = QTextCursor(block).currentTable()
        if table is not None:
            book.append(table_to_item(document, table))
            end = table.lastPosition()
            while block.isValid() and block.position() <= end:
                block = block.next()
            continue
        book.append(block_to_paragraph(block))
        block = block.next()
    # Leave out empty paragraphs at the very end.
    while book and book[-1]["type"] == "paragraph" and not "".join(r["text"] for r in book[-1]["runs"]).strip():
        book.pop()
    return book


# ----- PDF (made by Qt itself) ----------------------------------------------

def write_pdf(copy, file_name, title):
    width, height = TRIM_SIZES.get(page_settings["trim"], TRIM_SIZES[STANDARD_PAGE["trim"]])
    writer = QPdfWriter(file_name)
    writer.setTitle(title)
    writer.setCreator("Tabbed Writer")
    writer.setPageLayout(QPageLayout(
        QPageSize(QSizeF(width, height), QPageSize.Unit.Inch, "", QPageSize.SizeMatchPolicy.ExactMatch),
        QPageLayout.Orientation.Portrait,
        QMarginsF(page_settings["inside"], page_settings["top"], page_settings["outside"], page_settings["bottom"]),
        QPageLayout.Unit.Inch,
    ))
    text_width, text_height = page_text_size()
    copy.setDocumentMargin(0)
    copy.setPageSize(QSizeF(text_width, text_height))
    copy.print_(writer)


def word_support_missing():
    try:
        import docx  # noqa: F401
        return False
    except ImportError:
        return True


WORD_MISSING_MESSAGE = (
    "Word (.docx) files need the free python-docx add-on, which isn't installed.\n\n"
    "To add it, run this in a terminal:\n"
    "  ~/.local/share/tabbed-writer/.venv/bin/pip install python-docx"
)


# ---------------------------------------------------------------------------
# Saving copies (.docx, .epub, .pdf, .odt, .txt)
# ---------------------------------------------------------------------------

def is_centered_or_right(block_format):
    alignment = block_format.alignment()
    return bool(alignment & Qt.AlignmentFlag.AlignHCenter) or bool(alignment & Qt.AlignmentFlag.AlignRight)


def apply_book_paragraphs(document):
    # KDP: 0.2 inch first-line indent, no space before or after paragraphs.
    # Centered or right-aligned lines (like scene breaks) and tables are left alone.
    block = document.begin()
    while block.isValid():
        block_format = block.blockFormat()
        cursor = QTextCursor(block)
        if (
            block_format.headingLevel() == 0
            and cursor.currentTable() is None
            and block.text().strip()
            and not is_centered_or_right(block_format)
        ):
            book_paragraph = QTextBlockFormat()
            book_paragraph.setTextIndent(BOOK_INDENT)
            book_paragraph.setTopMargin(0)
            book_paragraph.setBottomMargin(0)
            cursor.mergeBlockFormat(book_paragraph)
        block = block.next()


def build_book_copy():
    current_snapshot()
    copy = QTextDocument()
    copy.setDefaultFont(QFont(DEFAULT_FONT, DEFAULT_SIZE))
    cursor = QTextCursor(copy)
    use_headings = sections.topLevelItemCount() > 1

    heading_text = QTextCharFormat()
    heading_text.setFontPointSize(18)
    heading_text.setFontWeight(QFont.Weight.Bold)

    ebook = document_kind == "ebook"
    for index in range(sections.topLevelItemCount()):
        item = sections.topLevelItem(index)
        if use_headings:
            heading_block = QTextBlockFormat()
            heading_block.setHeadingLevel(1)
            heading_block.setBottomMargin(24)
            if ebook:
                # KDP: each chapter title is a centered Heading 1 on a new page.
                heading_block.setAlignment(Qt.AlignmentFlag.AlignHCenter)
            if index == 0:
                cursor.setBlockFormat(heading_block)
            else:
                if ebook:
                    heading_block.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
                cursor.insertBlock(heading_block)
            cursor.insertText(item.text(0), heading_text)
            cursor.insertBlock(QTextBlockFormat(), QTextCharFormat())
        cursor.insertFragment(QTextDocumentFragment.fromHtml(item.data(0, CONTENT_ROLE) or ""))

    if ebook:
        apply_book_paragraphs(copy)
    return copy


def write_copy(file_name):
    target = Path(file_name)
    copy = build_book_copy()

    if target.suffix == ".txt":
        writer = QTextDocumentWriter(str(target))
        writer.setFormat(b"plaintext")
        if not writer.write(copy):
            QMessageBox.warning(window, "Save failed", "The file could not be written.")
            return
        window.statusBar().showMessage(f"Saved a copy as {target.name}", 4000)
        return

    if target.suffix == ".docx" and word_support_missing():
        QMessageBox.warning(window, "Save failed", WORD_MISSING_MESSAGE)
        return

    window.statusBar().showMessage(f"Saving {target.name}…")
    QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
    ebook = document_kind == "ebook"
    try:
        if target.suffix == ".pdf":
            write_pdf(copy, str(target), target.stem)
        elif target.suffix == ".docx":
            write_docx(document_to_book(copy), str(target), page_settings, TRIM_SIZES, target.stem, ebook)
        elif target.suffix == ".epub":
            write_epub(document_to_book(copy), str(target), target.stem, "", ebook)
        else:
            with tempfile.TemporaryDirectory() as work:
                odt = Path(work) / (target.stem + ".odt")
                writer = QTextDocumentWriter(str(odt))
                writer.setFormat(b"odf")
                if not writer.write(copy):
                    raise RuntimeError("The file could not be written.")
                finish_odt(odt, page_settings, target.stem)
                shutil.copyfile(odt, target)
    except Exception as problem:
        QApplication.restoreOverrideCursor()
        window.statusBar().clearMessage()
        QMessageBox.warning(window, "Save failed", str(problem))
        return
    QApplication.restoreOverrideCursor()
    window.statusBar().showMessage(f"Saved a copy as {target.name}", 4000)


def save_as_document():
    global DOCUMENT_FILE
    if DOCUMENT_FILE is None:
        start = WRITING_FOLDER / SUGGESTED_NAME
    else:
        start = DOCUMENT_FILE
    file_name, chosen = QFileDialog.getSaveFileName(
        window,
        "Save as",
        str(start),
        ";;".join(SAVE_TYPES),
    )
    if not file_name:
        return

    ending = SAVE_TYPES.get(chosen, ".twp")
    if not file_name.lower().endswith(ending):
        file_name += ending

    if ending == ".twp":
        DOCUMENT_FILE = Path(file_name)
        save_document()
        window.setWindowTitle(f"{DOCUMENT_FILE.name} - Tabbed Writer")
        return

    if sections.topLevelItemCount() > 1:
        QMessageBox.information(
            window,
            "Saving a copy",
            f"{ending} files can't keep tabs, so each tab's name becomes a heading in the copy"
            + (" (a centered chapter title on a new page)." if document_kind == "ebook" else ".")
            + "\n\nYour tabs stay as they are here.",
        )
    write_copy(file_name)


# ---------------------------------------------------------------------------
# Opening .twp, .odt, .docx and .txt
# ---------------------------------------------------------------------------

def html_from_plain_text(text):
    holder = QTextDocument()
    holder.setPlainText(text)
    return holder.toHtml()


def html_from_office_file(path):
    with tempfile.TemporaryDirectory() as work:
        libreoffice_convert(path, "html", work)
        result = Path(work) / (path.stem + ".html")
        if not result.exists():
            raise RuntimeError("LibreOffice couldn't read that file.")
        return result.read_text(encoding="utf-8", errors="replace")


def tabs_from_html(html):
    # Each Heading 1 (a chapter title) becomes its own tab.
    document = QTextDocument()
    document.setHtml(html)

    headings = []
    block = document.begin()
    while block.isValid():
        if (
            block.blockFormat().headingLevel() == 1
            and block.text().strip()
            and QTextCursor(block).currentTable() is None
        ):
            headings.append(block)
        block = block.next()

    if not headings:
        return [{"title": "Tab 1", "text": html}]

    def piece(start, end):
        if start >= end:
            return ""
        cursor = QTextCursor(document)
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        if not cursor.selectedText().strip():
            return ""
        return QTextDocumentFragment(cursor).toHtml()

    tabs = []
    opening = piece(0, headings[0].position() - 1)
    if opening:
        tabs.append({"title": "Front matter", "text": opening})
    last_position = document.characterCount() - 1
    for number, heading in enumerate(headings):
        start = heading.position() + heading.length()
        if number + 1 < len(headings):
            end = headings[number + 1].position() - 1
        else:
            end = last_position
        tabs.append({"title": heading.text().strip(), "text": piece(start, end)})
    return tabs


def open_document(file_name=None):
    global DOCUMENT_FILE, saved_snapshot, SUGGESTED_NAME, page_settings
    if not ok_to_discard():
        return
    if not isinstance(file_name, str):
        file_name, _ = QFileDialog.getOpenFileName(
            window,
            "Open",
            str(WRITING_FOLDER),
            "All supported (*.twp *.docx *.odt *.txt);;Tabbed Writer document (*.twp);;"
            "Word document (*.docx);;OpenDocument text (*.odt);;Plain text (*.txt)",
        )
    if not file_name:
        return
    path = Path(file_name)
    kind = path.suffix.lower()

    try:
        if kind == ".twp":
            load_document(path)
            DOCUMENT_FILE = path
        else:
            if kind == ".txt":
                tabs = [{"title": "Tab 1", "text": html_from_plain_text(
                    path.read_text(encoding="utf-8", errors="replace"))}]
            elif kind == ".docx":
                if word_support_missing():
                    raise RuntimeError(WORD_MISSING_MESSAGE)
                tabs = tabs_from_html(docx_to_html(str(path)))
            else:
                tabs = tabs_from_html(html_from_office_file(path))
            page_settings = dict(STANDARD_PAGE)
            set_document_kind("standard")
            apply_page_look()
            fill_tabs(tabs)
            DOCUMENT_FILE = None
            SUGGESTED_NAME = path.stem + ".twp"
    except Exception as problem:
        QMessageBox.warning(window, "Couldn't open file", str(problem))
        return

    window.setWindowTitle(f"{path.name} - Tabbed Writer")
    saved_snapshot = current_snapshot()


def new_document(kind="standard"):
    global DOCUMENT_FILE, saved_snapshot, SUGGESTED_NAME, page_settings
    if not ok_to_discard():
        return
    DOCUMENT_FILE = None
    SUGGESTED_NAME = "Untitled.twp"
    page_settings = dict(DEFAULT_PAGE if kind == "ebook" else STANDARD_PAGE)
    set_document_kind(kind)
    apply_page_look()
    first_title = "Chapter 1" if kind == "ebook" else "Tab 1"
    fill_tabs([{"title": first_title, "text": ""}])
    window.setWindowTitle("Untitled - Tabbed Writer")
    saved_snapshot = current_snapshot()


def toggle_ebook_formatting(turned_on):
    global document_kind
    document_kind = "ebook" if turned_on else "standard"
    use_book_paragraph_if_empty()
    update_counts()
    if turned_on:
        window.statusBar().showMessage(
            "eBook formatting on: saved copies follow the KDP and B&N rules", 5000
        )


# ---------------------------------------------------------------------------
# Page setup (for paperbacks)
# ---------------------------------------------------------------------------

def page_setup():
    global page_settings
    dialog = QDialog(window)
    dialog.setWindowTitle("Page setup (paperback)")
    form = QFormLayout(dialog)

    trim_box = QComboBox()
    trim_box.addItems(list(TRIM_SIZES))
    trim_box.setCurrentText(page_settings["trim"])
    form.addRow("Trim size:", trim_box)

    margin_boxes = {}
    for key, label in [("top", "Top"), ("bottom", "Bottom"),
                       ("inside", "Inside (gutter)"), ("outside", "Outside")]:
        box = QDoubleSpinBox()
        box.setRange(0.25, 2.0)
        box.setSingleStep(0.125)
        box.setDecimals(3)
        box.setSuffix(" in")
        box.setValue(float(page_settings[key]))
        margin_boxes[key] = box
        form.addRow(f"{label}:", box)

    form.addRow(QLabel(
        "KDP minimum inside margin by page count:\n"
        "  24–150 pages: 0.375 in      151–300: 0.5 in\n"
        "  301–500: 0.625 in      501–700: 0.75 in      701–828: 0.875 in\n"
        "Outside margin: at least 0.25 in.\n\n"
        "These settings set the page on screen and in PDF, Word and .odt copies.\n"
        "eBook readers ignore them, since readers choose their own layout."
    ))

    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    form.addRow(buttons)

    if not dialog.exec():
        return
    page_settings = {"trim": trim_box.currentText()}
    for key, box in margin_boxes.items():
        page_settings[key] = round(box.value(), 3)
    apply_page_look()
    update_counts()


# ---------------------------------------------------------------------------
# Insert: breaks and tables
# ---------------------------------------------------------------------------

def insert_page_break():
    cursor = editor.textCursor()
    break_format = cursor.blockFormat()
    break_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
    break_format.setHeadingLevel(0)
    cursor.insertBlock(break_format)
    editor.setTextCursor(cursor)
    window.statusBar().showMessage("Page break added (it shows in saved copies)", 4000)


def insert_line_break():
    editor.textCursor().insertText(" ")


def insert_table():
    dialog = QDialog(window)
    dialog.setWindowTitle("Insert table")
    form = QFormLayout(dialog)
    rows_box = QSpinBox()
    rows_box.setRange(1, 50)
    rows_box.setValue(3)
    columns_box = QSpinBox()
    columns_box.setRange(1, 10)
    columns_box.setValue(3)
    form.addRow("Rows:", rows_box)
    form.addRow("Columns:", columns_box)
    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    form.addRow(buttons)
    if not dialog.exec():
        return

    table_format = QTextTableFormat()
    table_format.setBorder(1)
    table_format.setBorderStyle(QTextFrameFormat.BorderStyle.BorderStyle_Solid)
    table_format.setBorderCollapse(True)
    table_format.setCellPadding(4)
    table_format.setCellSpacing(0)
    table_format.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
    editor.textCursor().insertTable(rows_box.value(), columns_box.value(), table_format)
    editor.setFocus()


def current_table_cell():
    cursor = editor.textCursor()
    table = cursor.currentTable()
    if table is None:
        return None, None
    return table, table.cellAt(cursor)


def add_table_row(offset):
    table, cell = current_table_cell()
    if table is not None:
        table.insertRows(cell.row() + offset, 1)


def add_table_column(offset):
    table, cell = current_table_cell()
    if table is not None:
        table.insertColumns(cell.column() + offset, 1)


def delete_table():
    table, _ = current_table_cell()
    if table is not None:
        table.removeRows(0, table.rows())


def delete_table_row():
    table, cell = current_table_cell()
    if table is None:
        return
    if table.rows() == 1:
        delete_table()
    else:
        table.removeRows(cell.row(), 1)


def delete_table_column():
    table, cell = current_table_cell()
    if table is None:
        return
    if table.columns() == 1:
        delete_table()
    else:
        table.removeColumns(cell.column(), 1)


def choose_color(title, start_color):
    color = QColorDialog.getColor(
        start_color, window, title, QColorDialog.ColorDialogOption.DontUseNativeDialog
    )
    return color if color.isValid() else None


def selected_table_cells():
    cursor = editor.textCursor()
    table = cursor.currentTable()
    if table is None:
        return None, []
    if cursor.hasComplexSelection():
        first_row, row_count, first_column, column_count = cursor.selectedTableCells()
    else:
        cell = table.cellAt(cursor)
        first_row, row_count, first_column, column_count = cell.row(), 1, cell.column(), 1
    cells = []
    for row in range(first_row, first_row + row_count):
        for column in range(first_column, first_column + column_count):
            cells.append(table.cellAt(row, column))
    return table, cells


def choose_cell_color():
    table, cells = selected_table_cells()
    if not cells:
        return
    color = choose_color("Cell color", QColor("#fff2cc"))
    if color is None:
        return
    for cell in cells:
        cell_format = cell.format()
        cell_format.setBackground(color)
        cell.setFormat(cell_format)


def clear_cell_color():
    table, cells = selected_table_cells()
    for cell in cells:
        cell_format = cell.format()
        cell_format.clearBackground()
        cell.setFormat(cell_format)


def choose_border_color():
    table, _ = selected_table_cells()
    if table is None:
        return
    table_format = table.format()
    color = choose_color("Border color", table_format.borderBrush().color())
    if color is None:
        return
    table_format.setBorderBrush(QBrush(color))
    table.setFormat(table_format)


def show_editor_menu(position):
    if not editor.textCursor().hasSelection():
        editor.setTextCursor(editor.cursorForPosition(position))
    menu = editor.createStandardContextMenu(position)
    if editor.textCursor().currentTable() is not None:
        menu.addSeparator()
        table_menu = menu.addMenu("Table")
        table_menu.addAction("Insert row above", lambda: add_table_row(0))
        table_menu.addAction("Insert row below", lambda: add_table_row(1))
        table_menu.addAction("Insert column left", lambda: add_table_column(0))
        table_menu.addAction("Insert column right", lambda: add_table_column(1))
        table_menu.addSeparator()
        table_menu.addAction("Delete row", delete_table_row)
        table_menu.addAction("Delete column", delete_table_column)
        table_menu.addAction("Delete table", delete_table)
        table_menu.addSeparator()
        table_menu.addAction("Cell color…", choose_cell_color)
        table_menu.addAction("Remove cell color", clear_cell_color)
        table_menu.addAction("Border color…", choose_border_color)
    menu.exec(editor.viewport().mapToGlobal(position))


# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def closeEvent(self, event):
        if ok_to_discard():
            event.accept()
        else:
            event.ignore()


class TabList(QTreeWidget):
    def dropEvent(self, event):
        if self.dropIndicatorPosition() == QAbstractItemView.DropIndicatorPosition.OnItem:
            event.ignore()
            return
        super().dropEvent(event)
        update_counts_soon()


app = QApplication(sys.argv)
app.setDesktopFileName("tabbed-writer")
for icon_file in (
    Path(__file__).with_name("tabbed-writer.svg"),
    Path.home() / ".local/share/icons/hicolor/scalable/apps/tabbed-writer.svg",
):
    if icon_file.exists():
        app.setWindowIcon(QIcon(str(icon_file)))
        break

# Documents written on one system show the matching fonts on the other.
for label in LINUX_FONTS:
    if ON_LINUX:
        QFont.insertSubstitution(OTHER_FONTS[label], LINUX_FONTS[label])
    else:
        QFont.insertSubstitution(LINUX_FONTS[label], OTHER_FONTS[label])

window = MainWindow()
window.setWindowTitle("Untitled - Tabbed Writer")

sections = TabList()
sections.setRootIsDecorated(False)
sections.setHeaderHidden(True)
sections.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)

editor = QTextEdit()
editor.setPlaceholderText("Type your document here...")
editor.document().setDefaultFont(QFont(DEFAULT_FONT, DEFAULT_SIZE))
editor.setObjectName("page")
editor.setFrameShape(QFrame.Shape.NoFrame)
editor.setMinimumWidth(420)
editor.setStyleSheet("#page { background: white; color: black; }")
apply_page_look()

page_area = QWidget()
page_area.setObjectName("pageArea")
page_area.setStyleSheet("#pageArea { background: #dadce0; }")
page_row = QHBoxLayout(page_area)
page_row.setContentsMargins(16, 16, 16, 0)
page_row.addStretch()
page_row.addWidget(editor, 100)
page_row.addStretch()
editor.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
editor.customContextMenuRequested.connect(show_editor_menu)

plus_button = QPushButton("+")
plus_button.setFixedWidth(30)
plus_button.setToolTip("Add tab (each tab is a chapter)")
plus_button.clicked.connect(add_main_tab)

close_tabs_button = QPushButton("←")
close_tabs_button.setFixedWidth(30)
close_tabs_button.setToolTip("Hide tabs")
close_tabs_button.clicked.connect(lambda: show_tabs_action.setChecked(False))

tabs_header = QHBoxLayout()
tabs_header.addWidget(close_tabs_button)
tabs_header.addWidget(QLabel("Document tabs"))
tabs_header.addStretch()
tabs_header.addWidget(plus_button)

word_panel = QWidget()
word_panel_layout = QVBoxLayout(word_panel)
word_panel_layout.setContentsMargins(0, 6, 0, 0)
word_panel_title = QLabel("Word count")
title_font = word_panel_title.font()
title_font.setBold(True)
word_panel_title.setFont(title_font)
panel_table = make_word_table()
panel_total = QLabel()
panel_total.setFont(title_font)
word_panel_layout.addWidget(word_panel_title)
word_panel_layout.addWidget(panel_table)
word_panel_layout.addWidget(panel_total)
word_panel.setVisible(False)

sidebar_layout = QVBoxLayout()
sidebar_layout.setContentsMargins(0, 0, 0, 0)
sidebar_layout.addLayout(tabs_header)
sidebar_layout.addWidget(sections, 3)
sidebar_layout.addWidget(word_panel, 2)

sidebar = QWidget()
sidebar.setLayout(sidebar_layout)

splitter = QSplitter(Qt.Orientation.Horizontal)
splitter.addWidget(sidebar)
splitter.addWidget(page_area)
splitter.setSizes([220, 780])

layout = QVBoxLayout()
layout.setContentsMargins(0, 0, 0, 0)
layout.addWidget(splitter)

container = QWidget()
container.setLayout(layout)
window.setCentralWidget(container)

count_label = QLabel()
window.statusBar().addPermanentWidget(count_label)
count_timer = QTimer()
count_timer.setSingleShot(True)
count_timer.setInterval(400)
count_timer.timeout.connect(update_counts)

first_tab = QTreeWidgetItem(sections, ["Tab 1"])
first_tab.setData(0, CONTENT_ROLE, "")
first_tab.setData(0, WORDS_ROLE, 0)

sections.currentItemChanged.connect(switch_section)
sections.setCurrentItem(first_tab)

saved_snapshot = current_snapshot()


def show_tab_menu(position):
    item = sections.itemAt(position)
    if item is None:
        return
    sections.setCurrentItem(item)
    tab_menu = QMenu(sections)
    tab_menu.addAction("Rename", rename_section)
    tab_menu.addSeparator()
    tab_menu.addAction("Delete", remove_section)
    tab_menu.exec(sections.viewport().mapToGlobal(position))


sections.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
sections.customContextMenuRequested.connect(show_tab_menu)
editor.textChanged.connect(update_counts_soon)
editor.cursorPositionChanged.connect(update_counts_soon)


# ---------------------------------------------------------------------------
# Text formatting
# ---------------------------------------------------------------------------

def toggle_bold():
    if editor.currentCharFormat().font().bold():
        editor.setFontWeight(QFont.Weight.Normal)
    else:
        editor.setFontWeight(QFont.Weight.Bold)


def toggle_italic():
    editor.setFontItalic(not editor.fontItalic())


def toggle_underline():
    editor.setFontUnderline(not editor.fontUnderline())


def set_alignment(alignment):
    editor.setAlignment(alignment)


def choose_text_color():
    color = choose_color("Text color", editor.textColor())
    if color is not None:
        editor.setTextColor(color)
    editor.setFocus()


def choose_highlight_color():
    color = choose_color("Highlight color", QColor("#fdf3c4"))
    if color is not None:
        highlight = QTextCharFormat()
        highlight.setBackground(color)
        editor.mergeCurrentCharFormat(highlight)
    editor.setFocus()


def remove_colors():
    cursor = editor.textCursor()
    if not cursor.hasSelection():
        plain = editor.currentCharFormat()
        plain.clearForeground()
        plain.clearBackground()
        editor.setCurrentCharFormat(plain)
        return
    start, end = cursor.selectionStart(), cursor.selectionEnd()
    piece = QTextCursor(editor.document())
    piece.beginEditBlock()
    position = start
    while position < end:
        piece.setPosition(position)
        piece.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
        plain = piece.charFormat()
        plain.clearForeground()
        plain.clearBackground()
        piece.setCharFormat(plain)
        position += 1
    piece.endEditBlock()
    editor.setFocus()


font_box = QComboBox()
font_box.addItems(list(FONT_CHOICES))
font_box.setToolTip("Font")

size_box = QComboBox()
size_box.addItems(SIZE_CHOICES)
size_box.setCurrentText(str(DEFAULT_SIZE))
size_box.setToolTip("Font size")


def set_font_choice(label):
    editor.setFontFamily(FONT_CHOICES[label])
    editor.setFocus()


def set_size_choice(size):
    editor.setFontPointSize(float(size))
    editor.setFocus()


def update_format_boxes(text_format):
    font_box.setCurrentText(FONT_LABELS.get(text_format.font().family(), "Serif"))
    size = text_format.fontPointSize()
    if size > 0:
        size_box.setCurrentText(str(int(size)))
    else:
        size_box.setCurrentText(str(DEFAULT_SIZE))


font_box.textActivated.connect(set_font_choice)
size_box.textActivated.connect(set_size_choice)
editor.currentCharFormatChanged.connect(update_format_boxes)


# ---------------------------------------------------------------------------
# Menus, with the formatting controls on the same line
# ---------------------------------------------------------------------------

def add_menu_item(menu, name, action_function, shortcut=None):
    action = QAction(name, window)
    if shortcut is not None:
        action.setShortcut(shortcut)
    action.triggered.connect(action_function)
    menu.addAction(action)
    return action


menu_bar = QMenuBar()

file_menu = menu_bar.addMenu("File")
add_menu_item(file_menu, "New", lambda: new_document("standard"), QKeySequence.StandardKey.New)
add_menu_item(file_menu, "New eBook", lambda: new_document("ebook"))
add_menu_item(file_menu, "Open…", open_document, QKeySequence.StandardKey.Open)
add_menu_item(file_menu, "Save", save_document, QKeySequence.StandardKey.Save)
add_menu_item(file_menu, "Save as…", save_as_document, QKeySequence.StandardKey.SaveAs)
file_menu.addSeparator()
add_menu_item(file_menu, "Page setup…", page_setup)
file_menu.addSeparator()
add_menu_item(file_menu, "Close", lambda: new_document("standard"), QKeySequence.StandardKey.Close)
add_menu_item(file_menu, "Exit", window.close)

edit_menu = menu_bar.addMenu("Edit")
add_menu_item(edit_menu, "Undo", editor.undo, QKeySequence.StandardKey.Undo)
add_menu_item(edit_menu, "Redo", editor.redo, QKeySequence.StandardKey.Redo)
edit_menu.addSeparator()
add_menu_item(edit_menu, "Cut", editor.cut, QKeySequence.StandardKey.Cut)
add_menu_item(edit_menu, "Copy", editor.copy, QKeySequence.StandardKey.Copy)
add_menu_item(edit_menu, "Paste", editor.paste, QKeySequence.StandardKey.Paste)
edit_menu.addSeparator()
add_menu_item(edit_menu, "Select all", editor.selectAll, QKeySequence.StandardKey.SelectAll)

view_menu = menu_bar.addMenu("View")
show_tabs_action = QAction("Show tabs", window)
show_tabs_action.setCheckable(True)
show_tabs_action.setChecked(True)
show_tabs_action.toggled.connect(sidebar.setVisible)
view_menu.addAction(show_tabs_action)


def set_word_count_panel(turned_on):
    word_panel.setVisible(turned_on)
    app_settings.setValue("show_word_count", turned_on)
    if turned_on:
        show_tabs_action.setChecked(True)
        update_counts()


show_word_count_action = QAction("Show word count", window)
show_word_count_action.setCheckable(True)
show_word_count_action.toggled.connect(set_word_count_panel)
view_menu.addAction(show_word_count_action)

insert_menu = menu_bar.addMenu("Insert")
add_menu_item(insert_menu, "Table…", insert_table)
break_menu = insert_menu.addMenu("Break")
add_menu_item(break_menu, "Page break", insert_page_break, QKeySequence("Ctrl+Return"))
add_menu_item(break_menu, "Line break\tShift+Enter", insert_line_break)

format_menu = menu_bar.addMenu("Format")
text_menu = format_menu.addMenu("Text")
bold_action = add_menu_item(text_menu, "Bold", toggle_bold, QKeySequence.StandardKey.Bold)
italic_action = add_menu_item(text_menu, "Italic", toggle_italic, QKeySequence.StandardKey.Italic)
underline_action = add_menu_item(text_menu, "Underline", toggle_underline, QKeySequence.StandardKey.Underline)
align_menu = format_menu.addMenu("Align")
add_menu_item(align_menu, "Left", lambda: set_alignment(Qt.AlignmentFlag.AlignLeft), QKeySequence("Ctrl+Shift+L"))
add_menu_item(align_menu, "Center", lambda: set_alignment(Qt.AlignmentFlag.AlignHCenter), QKeySequence("Ctrl+Shift+E"))
add_menu_item(align_menu, "Right", lambda: set_alignment(Qt.AlignmentFlag.AlignRight), QKeySequence("Ctrl+Shift+R"))
add_menu_item(align_menu, "Justified", lambda: set_alignment(Qt.AlignmentFlag.AlignJustify), QKeySequence("Ctrl+Shift+J"))
color_menu = format_menu.addMenu("Color")
text_color_action = add_menu_item(color_menu, "Text color…", choose_text_color)
highlight_action = add_menu_item(color_menu, "Highlight color…", choose_highlight_color)
add_menu_item(color_menu, "Remove colors", remove_colors)
format_menu.addSeparator()
ebook_action = QAction("eBook formatting (KDP / B&N)", window)
ebook_action.setCheckable(True)
ebook_action.toggled.connect(toggle_ebook_formatting)
format_menu.addAction(ebook_action)

tools_menu = menu_bar.addMenu("Tools")
add_menu_item(tools_menu, "Word count", show_word_count, QKeySequence("Ctrl+Shift+C"))
add_menu_item(tools_menu, "eBook words per page…", change_words_per_page)

top_bar = QWidget()
top_row = QHBoxLayout(top_bar)
top_row.setContentsMargins(0, 0, 6, 0)
top_row.setSpacing(4)
top_row.addWidget(menu_bar)
top_row.addSpacing(12)
top_row.addWidget(font_box)
top_row.addWidget(size_box)

for action, letter, style in [
    (bold_action, "B", "bold"),
    (italic_action, "I", "italic"),
    (underline_action, "U", "underline"),
]:
    action.setIconText(letter)
    action.setToolTip(action.text())
    button = QToolButton()
    button.setDefaultAction(action)
    button.setAutoRaise(True)
    button_font = QFont()
    if style == "bold":
        button_font.setBold(True)
    elif style == "italic":
        button_font.setItalic(True)
    else:
        button_font.setUnderline(True)
    button.setFont(button_font)
    top_row.addWidget(button)

for action, letter, look in [
    (text_color_action, "A", "QToolButton { color: #c62828; font-weight: bold; }"),
    (highlight_action, "A", "QToolButton { background: #fdf3c4; color: black; font-weight: bold; }"),
]:
    action.setIconText(letter)
    action.setToolTip(action.text().replace("…", ""))
    button = QToolButton()
    button.setDefaultAction(action)
    button.setAutoRaise(True)
    button.setStyleSheet(look)
    top_row.addWidget(button)

top_row.addStretch()
window.setMenuWidget(top_bar)

show_word_count_action.setChecked(
    str(app_settings.value("show_word_count", "false")).lower() == "true"
)
update_counts()
if len(sys.argv) > 1:
    open_document(sys.argv[1])
window.resize(1100, 700)
window.show()
editor.setFocus()

sys.exit(app.exec())
