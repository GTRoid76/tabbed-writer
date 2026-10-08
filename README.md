# Tabbed Writer

A simple word processor where **each tab is a chapter**. Write in tabs down the side, see your word counts as you go, and save your book as a Word file, EPUB or PDF that follows the formatting rules of **Amazon KDP** and **Barnes & Noble Press**.

Tabbed Writer is free and open source, and runs on Linux and Windows.

## Features

- **Tabs as chapters:** add tabs with **+**, rename or delete them with a right-click, and drag them to reorder.
- **Word count panel:** each tab's word count and the total, updated as you type (View → Show word count). Press **Ctrl+Shift+C** for a quick summary.
- **Page view:** you write on a white sheet sized to your page setup, with "Page X of Y" at the bottom.
- **Two kinds of document:**
  - **New:** an everyday document with ordinary word-processor settings.
  - **New eBook:** follows the KDP and B&N rules. Paragraphs get a 0.2 inch first-line indent with no extra spacing, and each tab becomes a centered "Heading 1" chapter title on a new page, which builds the eBook's table of contents.
- **Formatting:** serif, sans serif and monospace fonts (the same three styles Kindle Create offers), sizes, bold, italic, underline, alignment, text and highlight colors.
- **Tables:** insert a table, then right-click inside it to add or remove rows and columns, or change cell and border colors.
- **Page breaks and line breaks** (Ctrl+Enter and Shift+Enter).
- **Page setup:** KDP paperback trim sizes (6 x 9 in by default for eBooks, Letter for normal documents) with inside and outside margins.
- **Saving:**
  - **.twp** is Tabbed Writer's own format, the one that keeps your tabs.
  - **Word (.docx), EPUB, PDF, OpenDocument (.odt) and plain text (.txt)** are saved as copies, with each tab's name turned into a chapter heading.
- **Opening:** .twp, Word (.docx) and .txt files. Chapter headings in a Word file come in as separate tabs. Opening .odt files works too if LibreOffice is installed.
- Warns you before closing, opening or starting a new document with unsaved changes.

## Installing on Linux

You need Python 3, which comes with Linux Mint, Ubuntu and most other versions of Linux.

1. Download this project (on GitHub: **Code → Download ZIP**) and unzip it.
2. Open a terminal in the unzipped folder and run:

   ```bash
   bash install.sh
   ```

The installer puts the program in `~/.local/share/tabbed-writer`, adds **Tabbed Writer** to your Menu with its icon, and makes `.twp` files open in it when double-clicked. Your writing is kept in **Documents/Tabbed Writer**.

If the installer says Python couldn't create its private folder, run `sudo apt install python3-venv` and try again.

## Running on Windows

A ready-to-run Windows version is planned. Until then:

1. Install Python 3 from [python.org](https://www.python.org/downloads/) (tick **"Add python.exe to PATH"** during setup).
2. Open a Command Prompt in this project's folder and run:

   ```
   python -m pip install -r requirements.txt
   python main.py
   ```

## How it was made

Tabbed Writer was designed, directed and tested by **Shannon L. Walden**, with the code written with the help of **Claude**, an AI assistant made by Anthropic.

## License

Tabbed Writer is free software, released under the **GNU General Public License, version 3** (see the `LICENSE` file). You're welcome to use it, share it and change it. If you share a changed version, it must stay open source under the same license.
