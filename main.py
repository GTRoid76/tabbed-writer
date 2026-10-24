"""Tabbed Writer: a simple word processor where each tab is a chapter.

The book settings follow Amazon KDP's published guidance:
- eBooks: chapter titles use Heading 1 and are centered, body paragraphs
  have a 0.2 inch first-line indent with no extra space before or after,
  and Word (.docx) or EPUB files are uploaded.
- Paperbacks: standard trim sizes (6 x 9 in is the most common), mirrored
  inside/outside margins, and a print PDF.
"""

import io
import json
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from PySide6.QtCore import QSettings, QSizeF, Qt, QTimer
from PySide6.QtGui import (
    QAction,
    QBrush,
    QColor,
    QFont,
    QKeySequence,
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
CONTENT_ROLE = Qt.ItemDataRole.UserRole
WORDS_ROLE = Qt.ItemDataRole.UserRole + 1

# "standard" documents behave like any word processor.
# "ebook" documents follow the KDP / B&N Press formatting rules.
document_kind = "standard"

app_settings = QSettings("TabbedWriter", "TabbedWriter")
DEFAULT_WORDS_PER_PAGE = 250

# The same three kinds of font Kindle Create offers (serif, sans serif,
# monospace), using the free Liberation fonts that come with Linux Mint.
FONT_CHOICES = {
    "Serif": "Liberation Serif",
    "Sans serif": "Liberation Sans",
    "Monospace": "Liberation Mono",
}
DEFAULT_FONT = "Liberation Serif"
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
# LibreOffice helpers (used for .docx, .epub, .pdf and for opening files)
# ---------------------------------------------------------------------------

def libreoffice_convert(source, target_format, out_dir):
    office = shutil.which("soffice") or shutil.which("libreoffice")
    if office is None:
        raise RuntimeError("This needs LibreOffice, which wasn't found on this computer.")
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

    window.statusBar().showMessage(f"Saving {target.name}…")
    try:
        with tempfile.TemporaryDirectory() as work:
            odt = Path(work) / (target.stem + ".odt")
            writer = QTextDocumentWriter(str(odt))
            writer.setFormat(b"odf")
            if not writer.write(copy):
                raise RuntimeError("The file could not be written.")
            finish_odt(odt, page_settings, target.stem)

            if target.suffix == ".odt":
                result = odt
            else:
                out_dir = Path(work) / "out"
                out_dir.mkdir()
                libreoffice_convert(odt, target.suffix[1:], out_dir)
                result = out_dir / (target.stem + target.suffix)
                if not result.exists():
                    raise RuntimeError("LibreOffice couldn't make that kind of file.")
            shutil.copyfile(result, target)
    except Exception as problem:
        window.statusBar().clearMessage()
        QMessageBox.warning(window, "Save failed", str(problem))
        return
    window.statusBar().showMessage(f"Saved a copy as {target.name}", 4000)


def save_as_document():
    global DOCUMENT_FILE
    if DOCUMENT_FILE is None:
        start = Path.home() / "Documents" / SUGGESTED_NAME
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


def open_document():
    global DOCUMENT_FILE, saved_snapshot, SUGGESTED_NAME, page_settings
    if not ok_to_discard():
        return
    file_name, _ = QFileDialog.getOpenFileName(
        window,
        "Open",
        str(Path.home() / "Documents"),
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
    current_family = text_format.font().family()
    font_box.setCurrentText("Serif")
    for label, family in FONT_CHOICES.items():
        if family == current_family:
            font_box.setCurrentText(label)
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
window.resize(1100, 700)
window.show()
editor.setFocus()

sys.exit(app.exec())
