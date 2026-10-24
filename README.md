# Tabbed Writer

A simple word processor where **each tab is a chapter**. Write in tabs down the side, see your word counts as you go, and save your book as a Word file, EPUB or PDF that follows the formatting rules of **Amazon KDP** and **Barnes & Noble Press**.

Tabbed Writer is free and open source, and runs on Linux and Windows.

## Features

- **Tabs as chapters:** add tabs with **+**, rename or delete them with a right-click, and drag them to reorder.
- **Word count panel:** each tab's word count and the total, updated as you type (View → Show word count). Press **Ctrl+Shift+C** for a quick summary.
- **Page view:** you write on a white sheet sized to your page setup, with "Page X of Y" at the bottom.
- **Two kinds of document:**
  - **New:** an everyday document with ordinary word-processor settings, saved as .odt by default.
  - **New eBook:** a book project saved as .twp, following the KDP and B&N rules. Paragraphs get a 0.2 inch first-line indent with no extra spacing, and each tab becomes a centered "Heading 1" chapter title on a new page, which builds the eBook's table of contents.
- **Formatting:** paragraph styles (Normal text, Title, Subtitle, Headings 1–3), serif, sans serif and monospace fonts (the same three styles Kindle Create offers), sizes, bold, italic, underline, strikethrough, superscript and subscript, text and highlight colors, alignment, indents, line spacing, and clear formatting.
- **Lists:** bulleted and numbered lists; Tab and Shift+Tab move items in and out.
- **Spell check:** red underlines as you type, with suggestions, "Add to dictionary" and "Ignore" on right-click. Works offline.
- **Find and replace** (Ctrl+F, Ctrl+H), across one tab or all of them.
- **Pictures, links and horizontal lines** (Insert menu). Right-click a picture to resize it; Ctrl+click a link to open it.
- **Tables:** insert a table, then right-click inside it to add or remove rows and columns, or change cell and border colors.
- **Page breaks and line breaks** (Ctrl+Enter and Shift+Enter).
- **Printing** (Ctrl+P), with optional **page numbers** for printing, PDF and Word.
- **Page setup:** KDP paperback trim sizes (6 x 9 in by default for eBooks, Letter for normal documents) with inside and outside margins.
- **Saving:**
  - **Everyday documents** are saved as **OpenDocument (.odt)**, the open standard used by LibreOffice, or as **Word (.docx)**. Tabbed Writer quietly keeps your tabs inside its own .odt files, so they reopen exactly as you left them, while other programs simply see a normal document.
  - **eBooks** are saved as **.twp**, Tabbed Writer's book-project format.
  - **PDF, EPUB and plain text (.txt)** are saved as copies, with each tab's name turned into a heading.
- **Opening:** .twp, .odt, Word (.docx) and .txt files. Chapter headings in a Word file come in as separate tabs. Opening .odt files made by other programs uses LibreOffice, if it's installed.
- **Several documents at once:** File → New window (Ctrl+Shift+N) and File → Open in new window….
- **Never lose work:** documents save themselves every few minutes, unsaved work is recovered after a crash, and previous versions are kept (File → Version history).
- **Open recent**, **full screen** (F11), and a list of **keyboard shortcuts** (Ctrl+/).
- **Preferences:** choose your writing folder (for example, inside Dropbox), autosave timing, and how many previous versions to keep.
- Fully offline: nothing is ever sent over the internet.
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

## Installing on Windows

1. Go to the [Releases page](https://github.com/GTRoid76/tabbed-writer/releases/latest) and download **TabbedWriter.exe**.
2. Double-click it to run. There's nothing to install.

The first time, Windows may show a blue **"Windows protected your PC"** box, because the program isn't signed with a paid certificate. Click **More info**, then **Run anyway**. The first launch can take a few seconds while it unpacks.

Your writing is kept in **Documents\Tabbed Writer**.

## How it was made

Tabbed Writer was designed, directed and tested by **[your name]**, with the code written with the help of **Claude**, an AI assistant made by Anthropic.

## License

Tabbed Writer is free software, released under the **GNU General Public License, version 3** (see the `LICENSE` file). You're welcome to use it, share it and change it. If you share a changed version, it must stay open source under the same license.
