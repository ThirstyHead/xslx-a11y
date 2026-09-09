# xlsx-a11y

Audit and remediate Microsoft Excel `.xlsx` workbooks against **WCAG 2.1 AA** standards — standalone, deterministic, and headless.

`xlsx-a11y` evaluates SpreadsheetML / OOXML structures to uncover digital accessibility barriers across all four POUR principles (**Perceivable**, **Operable**, **Understandable**, **Robust**). It applies non-destructive, deterministic remediations and compiles comprehensive audit reports in **Markdown** (the single source of truth), **accessible HTML5** (styled via SMACSS), and **accessible tagged PDF**.

Part of the document accessibility quartet alongside [docx-a11y](https://github.com/ThirstyHead/docx-a11y), [pptx-a11y](https://github.com/ThirstyHead/pptx-a11y), and [pdf-a11y](https://github.com/ThirstyHead/pdf-a11y).

---

## Key Features

- **POUR Structure (WCAG 2.1 AA):** Every finding and recommendation is categorized under Perceivable, Operable, Understandable, or Robust, linking directly to canonical W3C Understanding documentation.
- **Social Model of Disability:** Language focuses strictly on workbook markup deficiencies and environmental barriers rather than individual limitations. Unit-tested language guards prevent medical-model or deficit phrasing.
- **Single Source of Truth:** Reports are authored as CommonMark Markdown, then compiled into accessible HTML5 and accessible tagged PDF with zero semantic or textual discrepancies.
- **SMACSS Theme Engine:** Build-free, text-based CSS theme architecture with 6 bundled palettes (`light`, `dark`, `ocean`, `forest`, `high-contrast`, `print`). Supports custom user themes in `~/.config/xlsx-a11y/themes/`.
- **Contrast Rigor (WCAG 1.4.3):** Evaluates cell-level text contrast against cell fill background colors (4.5:1 for normal text, 3:1 for large text >= 18pt or bold >= 14pt).
- **Interactive & Batch GUI:** Native cross-platform desktop application for bulk directory remediation, file drag-and-drop, real-time progress, and interactive barrier triage.
- **Tagged Accessible PDF:** Exports multi-page PDFs carrying `/Lang`, `/Title`, `/MarkInfo /Marked true`, and validated tag trees.
- **Deterministic & Structural Remediation:** Auto-remediates workbook title metadata, unmerges disruptive merged cells with value replication across the grid, converts raw data ranges to formal Excel Tables (`ListObject`) with headers enabled, and resets view focus to `A1`.
- **Interactive Triage:** Guided human-in-the-loop triage workflow for author-intent barriers: descriptive image alt text, decorative image markers, chart descriptions, and descriptive sheet tab names.
- **Strict Document Immutability:** Original files are never modified in-place; all remediations produce cryptographically verified new artifacts with SHA-256 provenance.

---

## Installation

### 1. Desktop GUI Application (Pre-built Installers)

Pre-built desktop installers with bundled dependencies (including PySide6/Qt) are available on the [GitHub Releases](https://github.com/ThirstyHead/xlsx-a11y/releases) page:

- **macOS (`.dmg`)**:
  1. Download `xlsx-a11y-v<version>-macos.dmg` (e.g. `xlsx-a11y-v0.1.0-macos.dmg`).
  2. Double-click to mount the disk image.
  3. Drag `xlsx-a11y.app` into `/Applications`.
  4. Launch `xlsx-a11y` from Spotlight, Launchpad, or the Applications folder.
  > **Note (macOS Gatekeeper)**: Because `xlsx-a11y` is an open-source binary distributed outside the Mac App Store without paid Apple Developer notarization, macOS blocks first launch with *“Apple could not verify xlsx-a11y is free of malware”*.
  > - **GUI bypass**: Right-click (or Control-click) `xlsx-a11y.app` in `/Applications`, select **Open**, and click **Open**. Alternatively, go to **System Settings > Privacy & Security**, scroll down to the **Security** section, and click **Open Anyway**.
  > - **Terminal bypass**: Run `xattr -cr /Applications/xlsx-a11y.app` (or `xattr -d com.apple.quarantine ~/Downloads/xlsx-a11y-*-macos.dmg` before opening the DMG).
- **Windows (`.exe`)**:
  1. Download `xlsx-a11y-setup-v<version>.exe` (e.g. `xlsx-a11y-setup-v0.1.0.exe`).
  2. Run the installer wizard to install into `Program Files` and create Start Menu / Desktop shortcuts.
- **Linux (`.AppImage`)**:
  1. Download `xlsx-a11y-v<version>-x86_64.AppImage` (e.g. `xlsx-a11y-v0.1.0-x86_64.AppImage`).
  2. Make it executable: `chmod +x xlsx-a11y-v*-x86_64.AppImage`.
  3. Run directly: `./xlsx-a11y-v*-x86_64.AppImage`.

### 2. Standalone Headless CLI Binary (No Python Required)

Single-file headless CLI executables are available on [GitHub Releases](https://github.com/ThirstyHead/xlsx-a11y/releases) for agentic workflows, CI/CD pipelines, and terminal environments:

```bash
# Example: Download macOS Apple Silicon standalone CLI binary
curl -LO https://github.com/ThirstyHead/xlsx-a11y/releases/latest/download/xlsx-a11y-cli-macos-arm64
chmod +x xlsx-a11y-cli-macos-arm64
sudo mv xlsx-a11y-cli-macos-arm64 /usr/local/bin/xlsx-a11y

# Verify installation
xlsx-a11y --help
```

### 3. Ephemeral Execution via `uvx`

Run `xlsx-a11y` instantly in any environment without managing Python virtual environments:

```bash
# Run headless CLI audit & remediation
uvx xlsx-a11y spreadsheet.xlsx --fix --format md,html,pdf,json

# Launch desktop GUI ephemerally
uvx --with "xlsx-a11y[gui]" xlsx-a11y-gui
```

### 4. Python Package via `pip` or `venv`

Requires **Python >= 3.10**.

#### Development Virtual Environment (`venv`)

```bash
# Clone and enter repository
git clone https://github.com/ThirstyHead/xlsx-a11y.git
cd xlsx-a11y

# Create and activate virtual environment
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\Activate.ps1

# Install with all dependencies
pip install -e ".[all]"
```

---

## Quickstart

### Example Test Documents

Two sample workbooks based on quarterly financial statements are provided in `examples/` for evaluating both the CLI and GUI:

- **`examples/Financial-Summary.xlsx` (Clean Reference):**
  A fully accessible workbook featuring formal Excel Tables with header rows, descriptive sheet tab names, high-contrast typography, descriptive chart titles, and workbook title metadata. Audits cleanly with **0 findings (100% WCAG 2.1 AA compliance)**.

- **`examples/Financial-Summary-test.xlsx` (Accessibility Barriers):**
  Intentionally includes digital accessibility barriers checked by `xlsx-a11y`: missing title metadata, default "Sheet1" tab name, merged cells disrupting grid navigation, raw tabular data without Table headers, missing chart title/alt text, and low contrast colors.

```bash
# Test the clean workbook (passes cleanly, exits 0):
xlsx-a11y "examples/Financial-Summary.xlsx" --format md,html,pdf,json --output-dir ./reports-clean

# Test the barrier workbook (detects barriers, exits 1):
xlsx-a11y "examples/Financial-Summary-test.xlsx" --format md,html,pdf,json --output-dir ./reports-barriers
```

Or drag both files directly into the desktop GUI (`xlsx-a11y --gui` or `xlsx-a11y-gui`) to test batch analysis, progress tracking, and interactive triage!

---

### 1. Audit a Workbook

Audit a workbook and generate all report formats (`.md`, `.html`, `.pdf`, `.json`):

```bash
xlsx-a11y workbook.xlsx --format md,html,pdf,json --output-dir ./reports
```

### 2. Remediate Deterministic Barriers

Automatically fix workbook title metadata, unmerge merged cells with value replication, convert rectangular tabular data into formal Excel Tables with headers, and reset view focus to `A1`:

```bash
xlsx-a11y workbook.xlsx --fix --out-xlsx workbook-remediated.xlsx
```

### 3. Interactive Triage

Launch the interactive human-in-the-loop CLI triage session to provide author-intent decisions (sheet names, alt text for images and charts):

```bash
xlsx-a11y workbook.xlsx --triage --out-xlsx workbook-triaged.xlsx
```

### 4. Batch Directory Auditing

Process an entire directory of workbooks in one command:

```bash
xlsx-a11y ./spreadsheets --batch --format md,html,pdf,json --output-dir ./batch-reports
```

---

## Audit Rules Reference

| Rule ID | WCAG SC | Severity | Fixable | Description |
|---|---|---|---|---|
| `title-missing` | 2.4.2 Level A | Critical | Yes | Workbook title is missing or blank in document core properties |
| `sheet-name-default` | 2.4.6 Level AA | Moderate | No (Triage) | Worksheet tab has default non-descriptive name (e.g. `Sheet1`) |
| `sheet-empty` | 1.3.1 Level A | Moderate | No | Worksheet contains no non-empty cells or data |
| `merged-cell` | 1.3.1 Level A | Serious | Yes | Merged cell ranges disrupt screen reader grid navigation and column associations |
| `table-header-missing` | 1.3.1 Level A | Serious | Yes | Tabular data is missing formal Excel Table declaration or table header row |
| `image-alt-missing` | 1.1.1 Level A | Serious | No (Triage) | Embedded drawing or image lacks descriptive alternative text |
| `chart-alt-missing` | 1.1.1 Level A | Serious | No (Triage) | Data chart lacks descriptive title or alternative text summary |
| `color-contrast` | 1.4.3 Level AA | Moderate | Yes | Cell font text does not meet WCAG contrast ratio (4.5:1 normal, 3:1 large/bold) |
| `link-text-vague` | 2.4.4 Level A | Moderate | No (Triage) | Hyperlink display text is vague (e.g. "click here") or a raw unformatted URL |

---

## Beyond Microsoft Excel's Built-In Accessibility Checker

Microsoft Excel includes a built-in accessibility checker (`Review` > `Check Accessibility`). While useful, `xlsx-a11y` enforces strict WCAG 2.1 AA criteria that native Office tools often miss or overlook:

1. **Merged Cells:**
   Microsoft Excel allows merged cells if they do not contain formulas. However, screen readers (JAWS, NVDA, VoiceOver) navigate cell grids linearly. Merged cells break coordinate tracking and header associations. `xlsx-a11y` unmerges cells and replicates data across the grid so assistive tech never encounters empty fields.
2. **Raw Tabular Ranges:**
   Excel's checker permits rectangular data ranges that visually appear as tables. Without a formal Excel `ListObject` (`Table`), screen readers cannot announce column headers or table boundaries during cell-by-cell navigation. `xlsx-a11y` converts rectangular data into structured Tables.
3. **Sheet Naming:**
   Excel often passes default sheet names (`Sheet1`, `Sheet2`) in template or older files. `xlsx-a11y` flags default sheet names to ensure screen reader users receive immediate context upon tab switching.

---

## License

MIT License. Developed with care for accessible digital documents.
