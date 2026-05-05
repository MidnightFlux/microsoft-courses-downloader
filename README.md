# Microsoft Courses Downloader

> *"Because clicking through 47 tiny units isn't learning—it's a workout for your mouse finger."*

---

## The Story (Why This Exists)

Picture this: It's 2 AM. Your Azure certification exam is in three days. You've got a course open on Microsoft Learn—let's say AI-102, because who doesn't love a good AI challenge at ungodly hours?

You start Module 1. Good. Clear. Then Module 2. Okay, getting longer. Then you realize something horrifying:

**This course has 12 modules.**  
**Each module has 5-8 units.**  
**Each unit is a separate page.**

Click. Read. Click. Read. Click. Read. Your browser now has 47 tabs open, your progress bar looks like a game of Snake, and you've lost track of where you were three times already. Did you finish Unit 3.4 or was that 4.3? Who knows? Not you.

And don't even get me started on trying to find that one specific paragraph you read yesterday. Was it in the Computer Vision module? Or was it NLP? *Where did it go?!*

The Microsoft Learn portal is great for bite-sized learning—if you're learning one bite at a time over a month. But when you need to actually *study*, to *review*, to have everything in one place you can search through? It's a maze designed by someone who really loves clicking buttons.

So I built this. Because:

- I was tired of playing "Find the Back Button"
- My browser shouldn't need therapy after a study session
- Ctrl+F is a basic human right when you're cramming for an exam
- Sometimes you just want a PDF you can read on a plane, offline, without 47 round-trips to the server

**This tool takes all those scattered micro-units and stitches them into beautiful, searchable, offline-friendly HTML and PDF files.** One file per module. All the content. Zero clicking.

You're welcome. Now go pass that exam.

---

## What It Does

- **Course Extraction**: Fetches learning paths and modules from any Microsoft Learn course URL
- **Catalog API Integration**: Uses the official Microsoft Learn Catalog API for accurate course structure
- **Content Processing**: Extracts and cleans main content from course pages
- **Smart Filtering**: Automatically skips interactive pages (knowledge checks, module assessments, exercises) that don't render usefully as static content
- **HTML Generation**: Creates beautifully formatted, combined HTML files for each module
- **PDF Conversion**: Converts HTML files to PDF using Playwright (optional)
- **Index Page**: Generates a clickable `index.html` in the course root with the full table of contents
- **Organised Output**: Numbered directories and files, Windows path-safe naming throughout

---

## Prerequisites

- **Python 3.10 or higher** — [Download Python](https://www.python.org/downloads/)

---

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/Laav/microsoft-courses-downloader.git
cd microsoft-courses-downloader
```

### 2. Set Up Virtual Environment (Recommended)

Creating a virtual environment keeps dependencies isolated from your system Python.

#### Windows (PowerShell)

```powershell
# Create virtual environment
python -m venv .venv

# Activate virtual environment
.venv\Scripts\Activate.ps1
```

> **Note**: If you get an execution policy error, run:
> `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

#### Windows (Command Prompt)

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

#### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies

With the virtual environment activated (you should see `(.venv)` in your prompt):

```bash
pip install -r requirements.txt
```

### 4. Install Playwright Browsers

Playwright requires browser binaries for PDF generation:

```bash
playwright install chromium
```

> **Linux users**: Run `playwright install-deps` first to install system dependencies.

---

## Usage

### Running the Script

1. **Activate your virtual environment** (if not already active):
   - **Windows**: `.venv\Scripts\Activate.ps1` (or `activate.bat`)
   - **macOS/Linux**: `source .venv/bin/activate`

2. **Run the script**:

```bash
python main.py
```

3. **Answer three prompts**:

```
Enter the Microsoft Learn course URL.
  Example: https://learn.microsoft.com/en-us/training/courses/az-140t00
> https://learn.microsoft.com/en-us/training/courses/az-104t00

Enter the output directory (press Enter to use current directory: C:\Temp\mcd):
>

Which output format do you want?
  1 = HTML only
  2 = PDF only  (HTML is generated as intermediate and then removed)
  3 = Both HTML and PDF  [default]
> 3
```

### Finding Course URLs

Browse all available Microsoft Learn courses at:  
**<https://learn.microsoft.com/en-us/training/browse/?resource_type=course>**

Copy any course URL and paste it when prompted. The URL is validated before processing starts.

---

## Example Output Structure

Given course `az-104t00` run from `C:\Temp\mcd` with output format **Both**:

```
C:\Temp\mcd\
└── az-104t00\
    ├── index.html                                          ← clickable table of contents
    ├── 01-az-104-prerequisites-azure-administrators\
    │   ├── 01-01-introduction-azure-cloud-shell.html
    │   ├── 01-01-introduction-azure-cloud-shell.pdf
    │   ├── 01-02-deploy-azure-infrastructure-json-arm-templates.html
    │   └── 01-02-deploy-azure-infrastructure-json-arm-templates.pdf
    ├── 02-az-104-manage-identities-governance-azure\
    │   ├── 02-01-understand-entra-id.html
    │   ├── 02-01-understand-entra-id.pdf
    │   └── ...
    └── ...
```

Key points:
- The **course directory** is named after the course code from the URL (e.g. `az-104t00`), not the full course title.
- **File names** are prefixed with `LP-Module` numbers (e.g. `06-02-` = learning path 6, module 2), making it easy to sort and group files across directories.
- **Interactive pages** (knowledge checks, module assessments, exercises) are automatically skipped and do not appear in the output.
- The **index.html** in the course root links to every generated HTML and PDF file.

---

## 🎓 Study Tips

Want to supercharge your learning? Upload the generated PDFs to **[Google NotebookLM](https://notebooklm.google.com)**!

NotebookLM is an AI-powered research assistant that can:

- 🎙️ **Generate Podcasts** — Create AI-generated audio summaries you can listen to on the go
- ❓ **Create Quizzes** — Generate practice questions to test your understanding
- 🃏 **Build Flashcards** — Make study cards for quick review sessions
- 💬 **Answer Questions** — Ask questions about the course content and get instant answers
- 🔗 **Connect Ideas** — Discover relationships between different concepts in the material

Simply upload the PDF files generated by this tool to NotebookLM, and start exploring your course content in new, interactive ways!

---

## How It Works

1. **URL Validation**: The course URL is validated before any network calls are made
2. **Catalog Fetching**: Course data is fetched from the Microsoft Learn Catalog API
3. **Learning Path Discovery**: All learning paths for the course are extracted
4. **Module Processing**: For each learning path, all modules are discovered
5. **Unit Extraction**: All unit links within each module are collected
6. **Smart Filtering**: Pages whose title matches a known interactive type (knowledge check, module assessment, exercise) are skipped automatically
7. **HTML Generation**: Remaining units are combined into a single styled HTML file per module
8. **PDF Conversion**: HTML files are optionally converted to PDF using Playwright
9. **Index Generation**: A `index.html` table of contents is written to the course root

---

## Architecture

The project uses an object-oriented design with clear separation of concerns:

| Class | Responsibility |
|---|---|
| `PathHelper` | Windows-safe path and filename construction |
| `HttpClient` | HTTP requests with consistent configuration |
| `CatalogService` | Interacts with Microsoft Learn Catalog API |
| `ContentService` | Fetches and processes web page content |
| `HtmlGenerator` | Generates combined HTML documents |
| `PdfGenerator` | Converts HTML to PDF using Playwright |
| `IndexGenerator` | Generates the course-level `index.html` |
| `CourseProcessor` | Orchestrates the entire extraction workflow |

---

## Configuration

Key constants can be modified at the top of `main.py`:

```python
# Pages skipped during HTML generation (interactive, no static value)
SKIP_PAGE_TITLE_PREFIXES = (
    "Knowledge check",
    "Module assessment",
    "Exercise - ",
)

# Maximum lengths for generated path components (Windows MAX_PATH safety)
MAX_COURSE_DIR_LENGTH = 32
MAX_LEARNING_PATH_DIR_LENGTH = 48
MAX_MODULE_FILE_STEM_LENGTH = 64
MAX_FULL_PATH_LENGTH = 235
```

---

## Requirements

- Python 3.10+
- `requests`
- `beautifulsoup4`
- `playwright`

---

## License

MIT License
