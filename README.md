# Microsoft Courses Downloader

A Python tool to extract learning content from Microsoft Learn courses and convert them to organized HTML and PDF files for offline reading.

## Features

- **Course Extraction**: Fetches learning paths and modules from any Microsoft Learn course URL
- **Catalog API Integration**: Uses the official Microsoft Learn Catalog API for accurate course structure
- **Content Processing**: Extracts and cleans main content from course pages
- **HTML Generation**: Creates beautifully formatted, combined HTML files for each module
- **PDF Conversion**: Converts HTML files to PDF using Playwright
- **Organized Output**: Generates structured output with numbered directories and files

## Installation

1. Clone or download this repository
2. Install Python dependencies:

```bash
pip install -r requirements.txt
```

3. Install Playwright browsers (required for PDF generation):

```bash
playwright install chromium
```

## Usage

Run the script:

```bash
python main.py
```

You'll be prompted to enter a Microsoft Learn course URL. Press Enter to use the default course (AI-102T00 - Designing and Implementing a Microsoft Azure AI Solution).

```
Enter the Microsoft Learn course URL (press Enter to use default: https://learn.microsoft.com/en-us/training/courses/ai-102t00):
> 
```

> **Finding Course URLs**: You can browse all available Microsoft Learn courses at [https://learn.microsoft.com/en-us/training/browse/?resource_type=course](https://learn.microsoft.com/en-us/training/browse/?resource_type=course)

### Example Output Structure

```
output/
├── 01-prepare-for-azure-ai-engineing/
│   ├── 01-introduction-to-azure-ai-studio.html
│   ├── 01-introduction-to-azure-ai-studio.pdf
│   ├── 02-understand-generative-ai.html
│   └── 02-understand-generative-ai.pdf
├── 02-computer-vision/
│   ├── 01-analyze-images.html
│   ├── 01-analyze-images.pdf
│   └── ...
└── ...
```

## How It Works

1. **Catalog Fetching**: The script fetches course data from the Microsoft Learn Catalog API
2. **Learning Path Discovery**: Extracts all learning paths associated with the course
3. **Module Processing**: For each learning path, discovers all modules
4. **Unit Extraction**: Fetches all unit links within each module
5. **HTML Generation**: Combines all units into a single, styled HTML file per module
6. **PDF Conversion**: Uses Playwright to convert HTML files to PDF format

## Architecture

The project uses an object-oriented design with clear separation of concerns:

| Class | Responsibility |
|-------|--------------|
| `HttpClient` | HTTP requests with consistent configuration |
| `CatalogService` | Interacts with Microsoft Learn Catalog API |
| `ContentService` | Fetches and processes web page content |
| `HtmlGenerator` | Generates combined HTML documents |
| `PdfGenerator` | Converts HTML to PDF using Playwright |
| `CourseProcessor` | Orchestrates the entire extraction workflow |

## Configuration

Default constants can be modified in `main.py`:

```python
DEFAULT_COURSE_URL = "https://learn.microsoft.com/en-us/training/courses/ai-102t00"
CATALOG_API_URL = "https://learn.microsoft.com/api/catalog/"
OUTPUT_BASE_DIR = "output"
```

## Requirements

- Python 3.8+
- requests
- beautifulsoup4
- playwright

## License

MIT License

