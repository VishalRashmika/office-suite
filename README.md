# Office Suite for Obsidian

<p align="center">
  <a href="https://obsidian.md"><img src="https://img.shields.io/badge/Obsidian-v1.4.0+-purple.svg?style=for-the-badge&logo=obsidian" alt="Obsidian Version" /></a>
  <a href="https://github.com"><img src="https://img.shields.io/badge/Version-v2.0-blue.svg?style=for-the-badge" alt="Release Version" /></a>
  <a href="https://obsidian.md/plugins?id=office-suite"><img src="https://img.shields.io/badge/dynamic/json?logo=obsidian&color=7C3AED&label=Downloads&query=%24%5B%27office-suite%27%5D.downloads&url=https%3A%2F%2Fraw.githubusercontent.com%2Fobsidianmd%2Fobsidian-releases%2Fmaster%2Fcommunity-plugin-stats.json&style=for-the-badge" alt="Obsidian Downloads" /></a>
  <a href="https://community.obsidian.md/plugins/office-suite"><img src="https://img.shields.io/badge/Community%20Plugins-Available-success.svg?style=for-the-badge&logo=obsidian" alt="Community Plugins" /></a>
  <a href="https://www.typescriptlang.org/"><img src="https://img.shields.io/badge/TypeScript-Strict-blue?style=for-the-badge&logo=typescript" alt="TypeScript" /></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-GPL--3.0-yellow.svg?style=for-the-badge" alt="License: GPL-3.0" /></a>
  <a href="https://github.com"><img src="https://img.shields.io/badge/PRs-Welcome-brightgreen.svg?style=for-the-badge" alt="PRs Welcome" /></a>
</p>

<p align="center">
  <a href="https://buymeacoffee.com/vishalrashmika">
    <img src="https://img.shields.io/badge/Buy%20Me%20A%20Coffee-Donate-orange.svg?style=for-the-badge&logo=buy-me-a-coffee" alt="Buy Me A Coffee" />
  </a>
</p>

---

> **View and edit Microsoft Word (.docx), Excel (.xlsx), PowerPoint (.pptx), CSV/TSV spreadsheets, and PDF documents directly inside Obsidian with native WYSIWYG editing, formatting, and formulas.**

**Office Suite** is a powerful, all-in-one productivity plugin for [Obsidian](https://obsidian.md) that brings a full office editing workflow into your knowledge vault. View, create, and edit Word documents, Excel spreadsheets, PowerPoint presentations, and PDF files **natively inside your Obsidian interface**.

> [!NOTE]
> **Desktop Only**: Office Suite requires Obsidian Desktop (Windows, macOS, Linux) to operate native document parsers and binary manipulation engines.

No external applications, no converting your files to Markdown, and no loss of formatting. Edit documents, spreadsheets, presentations, and PDFs side-by-side with your notes, with full autosave round-tripped directly back to the original binary files.

---

## Demos & Walkthroughs

### 1. Word Document Editing (.docx)
<p align="center">
  <img src="https://i.imgur.com/YHCt65g.gif" alt="Office Suite Word Editor Demo" />
</p>

### 2. PDF Editing & Highlighting (.pdf)
<p align="center">
  <video src="https://res.cloudinary.com/de7qgm8tw/video/upload/v1794228363/pdf-editor_qtafjl.mp4" controls width="100%">
    <p>Your browser does not support the video tag. <a href="https://res.cloudinary.com/de7qgm8tw/video/upload/v1794228363/pdf-editor_qtafjl.mp4">Watch PDF Editing Demo</a></p>
  </video>
  <br>
  <em>Video 1: Native PDF text selection, Edge-style floating highlight menu, ink pen drawing, and text notes.</em>
</p>

### 3. PowerPoint Presentation Editing & Slideshow (.pptx)
<p align="center">
  <video src="https://res.cloudinary.com/de7qgm8tw/video/upload/v1794228164/pptx-editor_c1wpsf.mp4" controls width="100%">
    <p>Your browser does not support the video tag. <a href="https://res.cloudinary.com/de7qgm8tw/video/upload/v1794228164/pptx-editor_c1wpsf.mp4">Watch PowerPoint Editing Demo</a></p>
  </video>
  <br>
  <em>Video 2: PowerPoint slide deck editing, text and shape creation, speaker notes drawer, and fullscreen slideshow.</em>
</p>

---

## Key Highlights

- **Native Word (.docx) Editor**: Full WYSIWYG rich-text editing with true OOXML lossless round-tripping.
- **Native Spreadsheet (.xlsx, .csv, .tsv) Editor**: Interactive multi-sheet grid interface with formula engine, cell formatting, and number styles.
- **Native PowerPoint (.pptx) Editor & Slideshow**: WYSIWYG slide editing, text box & vector shape creation, speaker notes drawer, slide lifecycle management, and fullscreen presentation mode with transitions.
- **Native PDF (.pdf) Editor & Highlighter**: In-DOM PDF engine with Mozilla PDF.js precision text layer, Edge-style floating quick menu, multiplied vibrant highlights, smooth freehand ink drawing, text notes, and lossless `pdf-lib` saving.
- **Adaptive Theme Support**: Seamlessly toggle between Obsidian Dark Paper mode and true White Paper mode for optimal reading and writing comfort.
- **Instant Debounced Autosave**: Changes write straight to your vault files on disk without interrupting your writing flow.
- **Bidirectional Deep Linking**: Deep link directly to specific slides (`[[deck.pptx#slide=3]]`) or PDF pages (`[[doc.pdf#page=2]]`).
- **100% Local & Privacy-First**: Everything runs entirely on your local machine. No cloud telemetry, no remote servers, and no external dependencies.

---

## Comprehensive Feature Matrix

### 1. Word Document Editor (.docx)

Edit standard `.docx` files generated by Microsoft Word, LibreOffice, or Google Docs with true round-trip fidelity.

<p align="center">
  <img src="https://i.imgur.com/UhRsVjY.png" alt="Word Document Editor (.docx) Preview" />
</p>

| Category | Supported Features |
| :--- | :--- |
| **Typography & Fonts** | Custom font families (*Calibri, Arial, Times New Roman, Georgia, Courier New, Verdana, Segoe UI, Trebuchet MS, Tahoma*), font sizes (*8pt – 72pt*), Clear Formatting |
| **Rich Text Styling** | **Bold**, *Italic*, <u>Underline</u>, ~~Strikethrough~~, Subscript (x₂), Superscript (x²), `Inline Code`, Blockquotes |
| **Colors** | Text color picker with custom HEX selector, Background highlight color picker |
| **Headings & Structure** | Full heading hierarchy (**H1 – H6**) mapped 1:1 to Word styles, Paragraphs, Horizontal Dividers, Page Breaks (`Ctrl+Enter`) |
| **Paragraph & Alignment** | Left, Center, Right, and Justified alignment; customizable line spacing (*1.0, 1.15, 1.5, 2.0*); Indent & Outdent controls |
| **Lists** | Bulleted lists, Numbered / Ordered lists, Nested list indentation and lift controls |
| **Tables** | Insert custom tables, Add Row Above / Below, Add Column Left / Right, Delete Row / Column / Table, Custom Cell Shading (Blue, Gray, Yellow, Green, Clear) |
| **Images & Media** | Embedded document images, interactive resize handles with aspect-ratio locking, alignment presets (Left, Center, Right), dimension tooltips |
| **Preservation Engine** | Unedited OOXML elements (headers, footers, `styles.xml`, `numbering.xml`, themes, and page margins) are preserved intact |

---

### 2. Spreadsheet & Data Editor (.xlsx, .csv, .tsv)

View, edit, calculate, and format Excel spreadsheets and tabular files seamlessly inside your notes.

<p align="center">
  <img src="https://i.imgur.com/IZbjhfI.png" alt="Spreadsheet & Data Editor (.xlsx) Preview" />
</p>

| Category | Supported Features |
| :--- | :--- |
| **Grid Engine** | Smooth spreadsheet grid with row/column headers, active cell cursor, range selection, and keyboard navigation |
| **Formula Engine** | Real-time calculation with `=SUM()`, `=AVERAGE()`, `=COUNT()`, `=MAX()`, `=MIN()`, `=IF()`, `=CONCAT()`, `=ROUND()`, and algebraic expressions |
| **Formula Bar** | Dedicated formula bar showing current cell coordinate (e.g., `A1`, `C12`) with full interactive formula editing |
| **Number Formatting** | Quick Currency format (`$`), Percentage format (`%`), Increase decimal places (`.00`), Decrease decimal places (`.0`) |
| **Cell Typography** | Font family selector, font size controls, **Bold**, *Italic*, <u>Underline</u>, ~~Strikethrough~~ |
| **Cell Styling & Colors**| Text color palette, Cell background / fill color picker, Cell border toggling |
| **Cell Alignment** | Left align (text default), Center align, Right align (numeric default) |
| **Row & Column Controls**| Insert row below, Insert column right, Delete selected row, Delete selected column |
| **Multi-Sheet Workbooks**| Workbook tab bar: switch sheets, create new sheet (`+`), and double-click to rename sheet tabs |
| **CSV & TSV Support** | Native parsing and serialization of comma-separated and tab-separated files with full grid capabilities |
| **History Stack** | Full undo (`Ctrl+Z`) and redo (`Ctrl+Y`) history stack for spreadsheet edits |

---

### 3. PowerPoint Presentation & Slideshow (.pptx)

View, edit, present, and organize Microsoft PowerPoint decks natively inside your vault.

<p align="center">
  <!-- PLACEHOLDER FOR PPTX PREVIEW IMAGE / GIF -->
  <img src="https://i.imgur.com/IZbjhfI.png" alt="PowerPoint Presentation Editor (.pptx) Preview" />
</p>

| Category | Supported Features |
| :--- | :--- |
| **WYSIWYG Slide Editing** | Direct in-place editing of text boxes, slide titles, body paragraphs, and formatting |
| **Shapes & Elements** | Add new Text Boxes and Vector Shapes (Rectangles, Rounded Rectangles, Ellipses, Triangles, Arrows, Lines) with solid/gradient fills and borders |
| **Speaker Notes Drawer** | Dedicated collapsible speaker notes drawer with real-time editing and slide synchronization |
| **Slide Lifecycle CRUD** | Add new blank slides, duplicate existing slides, delete slides, and reorder slide sequence |
| **Transitions & Animations** | Hardware-accelerated slide transitions: Fade, Push/Slide, Wipe, Zoom, and 3D Flip |
| **Fullscreen Slideshow** | Distraction-free presentation mode (`F5`) with presentation timer, slide counter, keyboard navigation, and auto-hiding HUD |
| **Slide Sorter & Navigation** | Slide thumbnails list, jump-to-slide navigation, and previous/next controls |
| **1-Click Markdown Export** | Export presentation outlines, slide contents, bullet points, and speaker notes directly into a structured Obsidian Markdown note |
| **Slide Deep Linking** | Bidirectional Obsidian links (`[[presentation.pptx#slide=3]]`) that open directly to the target slide |
| **OOXML Compliance** | Clean ECMA-376 XML round-tripping preserving themes, masters, and slide layouts |

---

### 4. Native PDF Editor, Highlighting & Annotation Engine (.pdf)

View, annotate, highlight, and edit PDF documents natively with 100% self-contained in-DOM rendering (no external apps, zero iframes, no 404s).

<p align="center">
  <!-- PLACEHOLDER FOR PDF PREVIEW IMAGE / GIF -->
  <img src="https://i.imgur.com/UhRsVjY.png" alt="PDF Editor & Highlighting Preview" />
</p>

| Category | Supported Features |
| :--- | :--- |
| **Pixel-Perfect Text Layer** | Mozilla PDF.js native text layer with exact `--scale-factor` coordinate mapping. Native multi-column selection with zero coordinate drift |
| **Edge-Style Floating Menu** | Selecting any text displays a sleek floating pill menu with 1-click color highlighters, Copy (`Ctrl+C`), and Add Note |
| **Multiplied Blending** | Highlights render with `mix-blend-mode: multiply` so the underlying text remains deep black, crisp, and completely legible |
| **Instant Highlight Tool** | Dedicated Highlight pen tool in top toolbar: simply drag across text to immediately highlight upon mouse release |
| **Highlight Editing Popover** | Click any existing highlight to change its color on the fly or delete it with a single click |
| **Freehand Drawing (Ink)** | Smooth vector pen with thickness presets (*2px Thin, 4px Medium, 8px Thick*) and full color palette |
| **Text Callout Notes** | Click anywhere on a page to insert editable text callout notes |
| **Eraser & Undo** | Eraser tool deletes any highlight, ink stroke, or note on click; full Undo (`Ctrl+Z` / `Cmd+Z`) support |
| **Lossless Saving (pdf-lib)** | Save modifications directly back to the PDF binary in your vault (`Ctrl+S` / `Cmd+S`) without flattening or quality loss |
| **Page Navigation & Zoom** | Page input jump, Next/Prev buttons, 90° Clockwise Rotation, Zoom In/Out, and Fit-to-Width |
| **Annotations to Markdown** | 1-click export of highlights and notes to Markdown with Obsidian deep links (`[[document.pdf#page=2]]`) |

---

## Installation

### Method 1: Obsidian Community Plugins (Recommended)

Once listed in the official Obsidian Community Plugins directory:

1. Open **Obsidian Settings**.
2. Navigate to **Community plugins** and ensure **Restricted mode** is turned off.
3. Click **Browse** and search for **"Office Suite"**.
4. Click **Install**, then click **Enable**.
5. Open any `.docx`, `.xlsx`, `.pptx`, `.csv`, `.tsv`, or `.pdf` file from your file explorer to start working!

---

### Method 2: Beta Testing via Obsidian BRAT

You can test the latest beta releases using the [Obsidian BRAT](https://github.com/TfTHacker/obsidian42-brat) plugin:

1. Install **BRAT** from Community Plugins.
2. Open **Settings → BRAT → Add Beta plugin**.
3. Paste the repository URL: `https://github.com/vishalrashmika/obsidian-office-suite`.
4. Click **Add Plugin** and enable **Office Suite**.

---

### Method 3: Manual Installation

1. Download `manifest.json`, `main.js`, and `styles.css` from the latest release.
2. Open your Obsidian Vault folder in your file manager.
3. Navigate to `.obsidian/plugins/` (create the folder if it doesn't exist).
4. Create a new folder named `office-suite`.
5. Copy `manifest.json`, `main.js`, and `styles.css` into that folder.
6. In Obsidian, go to **Settings → Community plugins** and enable **Office Suite**.

---

## Quick Start Guide

1. **Opening Any Office File**: Click on any `.docx`, `.xlsx`, `.pptx`, `.csv`, `.tsv`, or `.pdf` file in your Obsidian vault. It opens instantly inside an Office Suite tab.
2. **Editing Word Docs**: Type directly in the WYSIWYG editor. Use the top toolbar to style typography, headings, tables, and images.
3. **Working with Spreadsheets**: Select cells, input data, and use formulas like `=SUM(A1:A10)`. Manage multiple sheets using the bottom workbook tab bar.
4. **Editing & Presenting Slide Decks**: Edit slide text, add shapes, add speaker notes in the bottom drawer, or click **Slideshow** for full-screen presentation mode.
5. **Annotating & Highlighting PDFs**: Select text to reveal the floating highlight menu, or use the top toolbar's Highlight, Ink Pen, or Text Note tools. Press `Ctrl+S` to save changes losslessly to the PDF.
6. **Autosaving**: Edits to `.docx`, `.xlsx`, `.pptx`, and `.csv` files autosave automatically with a debounced background worker. PDFs can be saved at any time with `Ctrl+S` or the Save button.

---

## Shortcuts and Hotkeys

| Action | Shortcut | Target Formats |
| :--- | :--- | :--- |
| **Save to Vault** | `Ctrl + S` / `Cmd + S` | `.docx`, `.xlsx`, `.pptx`, `.pdf` |
| **Undo** | `Ctrl + Z` / `Cmd + Z` | All formats |
| **Redo** | `Ctrl + Y` / `Cmd + Shift + Z` | `.docx`, `.xlsx`, `.pptx` |
| **Bold** | `Ctrl + B` / `Cmd + B` | `.docx`, `.xlsx`, `.pptx` |
| **Italic** | `Ctrl + I` / `Cmd + I` | `.docx`, `.xlsx`, `.pptx` |
| **Underline** | `Ctrl + U` / `Cmd + U` | `.docx`, `.xlsx`, `.pptx` |
| **Page Break** | `Ctrl + Enter` / `Cmd + Enter` | `.docx` |
| **Edit Cell** | `F2` | `.xlsx`, `.csv`, `.tsv` |
| **Start Slideshow** | `F5` / Command Palette | `.pptx` |
| **Next Slide** | `Right Arrow` / `Space` / `PageDown` | `.pptx` |
| **Previous Slide** | `Left Arrow` / `Backspace` / `PageUp` | `.pptx` |
| **Exit Slideshow** | `Escape` | `.pptx` |
| **Dismiss Floating Menu** | `Escape` | `.pdf` |

---

## Contributing

Contributions, feature requests, and bug reports are welcome! Help us make Office Suite the premier document and productivity tool for the Obsidian ecosystem.

### How to Contribute

1. **Report a Bug / Request a Feature**: Open an Issue on GitHub with detailed reproduction steps or feature descriptions.
2. **Submit a Pull Request**:
   - Fork the repository.
   - Create a feature branch (`git checkout -b feature/amazing-feature`).
   - Commit your changes (`git commit -m 'Add amazing feature'`).
   - Push to your branch (`git push origin feature/amazing-feature`).
   - Open a Pull Request.

### Development Setup

```bash
# Clone the repository
git clone https://github.com/vishalrashmika/obsidian-office-suite.git
cd obsidian-office-suite

# Install dependencies
npm install

# Run static analysis and TypeScript check
npm run check

# Run comprehensive test suite
npm test

# Start development build with watch mode
npm run dev

# Build production bundle
npm run build
```

---

## Support the Project

If you find **Office Suite** valuable for your daily note-taking and workflow, please consider supporting its continuous development:

<p align="center">
  <a href="https://buymeacoffee.com/vishalrashmika" target="_blank">
    <img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" height="50" width="210">
  </a>
</p>

Your support helps cover development tools, test devices, and future feature expansions! Don't forget to star the repository on GitHub!

---

## Author

Created by [Vishal Rashmika](https://vishalrashmika.com).

---

## License

This project is licensed under the [GPL-3.0 License](LICENSE).
