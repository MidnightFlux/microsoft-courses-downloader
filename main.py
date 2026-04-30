"""
Script to extract learning path links from a Microsoft Learn course page.
Uses the Microsoft Learn Catalog API.

This version shortens directory and file names so it behaves better on Windows
systems that still hit MAX_PATH limitations.

Output behavior:
- The script asks for the Microsoft Learn course URL.
- The script asks for the output base directory.
- The script asks which output format is desired (HTML, PDF, or both).
- The course directory is created directly inside the selected output directory.
- An index.html is generated in the course root directory.

Example:
If this script is run from:
    C:\\Temp\\mcd

Course URL:
    https://learn.microsoft.com/en-us/training/courses/az-140t00

Output directory:
    Press Enter for default

Then files are written to:
    C:\\Temp\\mcd\\az-140t00\\
"""

import html
import os
import re
import requests
import asyncio
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from playwright.async_api import async_playwright


# =============================================================================
# Constants
# =============================================================================

CATALOG_API_URL = "https://learn.microsoft.com/api/catalog/"
LEARN_BASE_URL = "https://learn.microsoft.com"
LEARN_COURSE_PATH_PREFIX = "/training/courses/"

# Default output base directory.
# "." means: current working directory.
DEFAULT_OUTPUT_BASE_DIR = "."

REQUEST_TIMEOUT = 30
CATALOG_TIMEOUT = 60

MAX_COURSE_DIR_LENGTH = 32
MAX_LEARNING_PATH_DIR_LENGTH = 48
MAX_MODULE_FILE_STEM_LENGTH = 64
MAX_FULL_PATH_LENGTH = 235

DEFAULT_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.0"
}

# Valid output format choices.
OUTPUT_FORMAT_HTML = "html"
OUTPUT_FORMAT_PDF = "pdf"
OUTPUT_FORMAT_BOTH = "both"

HTML_STYLES = """
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; max-width: 900px; margin: 0 auto; padding: 20px; line-height: 1.6; }
    h1 { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
    h2 { color: #333; margin-top: 40px; border-bottom: 1px solid #ddd; padding-bottom: 8px; }
    .section { margin-bottom: 40px; }
    .section-header { background: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
    .section-header a { color: #0078d4; text-decoration: none; }
    .section-header a:hover { text-decoration: underline; }
    img { max-width: 100%; height: auto; }
    pre { background: #f4f4f4; padding: 15px; overflow-x: auto; border-radius: 5px; }
    code { background: #f4f4f4; padding: 2px 5px; border-radius: 3px; }
    table { border-collapse: collapse; width: 100%; margin: 15px 0; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background: #f5f5f5; }
    .NOTE, .TIP { padding: 12px 15px; margin: 15px 0; border-radius: 5px; border-left: 4px solid; }
    .NOTE { background-color: #e7f3ff; border-color: #0078d4; }
    .NOTE > p:first-child { font-weight: bold; color: #0078d4; margin-top: 0; }
    .TIP { background-color: #e8f5e9; border-color: #4caf50; }
    .TIP > p:first-child { font-weight: bold; color: #2e7d32; margin-top: 0; }
"""

INDEX_STYLES = """
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; max-width: 960px; margin: 0 auto; padding: 30px 20px; line-height: 1.6; background: #fafafa; }
    h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 12px; margin-bottom: 8px; }
    .course-url { color: #555; font-size: 0.9em; margin-bottom: 30px; }
    .course-url a { color: #0078d4; text-decoration: none; }
    .course-url a:hover { text-decoration: underline; }
    .lp-block { background: #fff; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 24px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,.06); }
    .lp-header { background: #0078d4; color: #fff; padding: 12px 18px; display: flex; align-items: center; gap: 10px; }
    .lp-header .lp-num { background: rgba(255,255,255,.25); border-radius: 4px; padding: 1px 8px; font-size: 0.85em; font-weight: bold; }
    .lp-header a { color: #fff; text-decoration: none; font-weight: 600; }
    .lp-header a:hover { text-decoration: underline; }
    .module-list { list-style: none; margin: 0; padding: 0; }
    .module-list li { border-bottom: 1px solid #f0f0f0; padding: 10px 18px; display: flex; align-items: center; gap: 12px; }
    .module-list li:last-child { border-bottom: none; }
    .module-num { color: #888; font-size: 0.82em; font-weight: bold; min-width: 48px; }
    .module-title { flex: 1; font-weight: 500; }
    .file-links { display: flex; gap: 8px; }
    .file-links a { display: inline-block; padding: 3px 10px; border-radius: 4px; font-size: 0.82em; text-decoration: none; font-weight: 600; }
    .link-html { background: #e7f3ff; color: #0078d4; border: 1px solid #b3d4f7; }
    .link-html:hover { background: #cce4ff; }
    .link-pdf  { background: #fdecea; color: #c62828; border: 1px solid #f5bcb8; }
    .link-pdf:hover  { background: #fbc8c5; }
    .footer { margin-top: 30px; font-size: 0.8em; color: #aaa; text-align: center; }
"""

# Generic phrase replacements for slug generation.
# Keep this list minimal and course-agnostic.
# Do NOT add product- or course-specific terms here.
COMMON_PHRASE_REPLACEMENTS = {
    "microsoft learn": "",
    "microsoft": "",
}

STOPWORDS = {
    "a", "an", "and", "by", "for", "from", "in", "into",
    "of", "on", "or", "the", "to", "using", "with",
}


# =============================================================================
# Data Classes
# =============================================================================


@dataclass
class PageContent:
    """Represents extracted content from a web page."""
    title: str
    content: str
    url: str


@dataclass
class ModuleIndexEntry:
    """Holds index information for one processed module."""
    lp_index: int
    module_index: int
    module_title: str
    module_url: str
    html_file: Optional[str]   # Absolute path, or None if not kept
    pdf_file: Optional[str]    # Absolute path, or None if not generated


@dataclass
class LearningPathIndexEntry:
    """Holds index information for one learning path."""
    lp_index: int
    lp_title: str
    lp_url: str
    lp_dir: str                # Absolute path to LP directory
    modules: list[ModuleIndexEntry] = field(default_factory=list)


# =============================================================================
# Path Helpers
# =============================================================================


class PathHelper:
    """Helpers for creating short, Windows-friendly path names."""

    @staticmethod
    def sanitize_component(name: str) -> str:
        """Sanitize a path component for Windows filesystems."""
        name = re.sub(r'[<>:"/\\|?*]', "-", name)
        name = re.sub(r"\s+", " ", name).strip(" .")
        return name or "untitled"

    @staticmethod
    def slugify(text: str) -> str:
        """Convert a string into a compact path-friendly slug."""
        text = text.lower().strip()
        for source, target in COMMON_PHRASE_REPLACEMENTS.items():
            text = text.replace(source, target)
        text = text.replace("&", " and ")
        text = re.sub(r"[^a-z0-9]+", " ", text)
        words = [w for w in text.split() if w and w not in STOPWORDS]
        slug = "-".join(words)
        return re.sub(r"-+", "-", slug).strip("-")

    @staticmethod
    def shorten_title(
        title: str,
        fallback: str,
        max_length: int,
        prefix: str = "",
    ) -> str:
        """Create a short filename/directory component from a title."""
        slug = PathHelper.slugify(title)
        fallback_slug = PathHelper.slugify(fallback)
        base = slug or fallback_slug or "untitled"

        if prefix:
            prefix = PathHelper.sanitize_component(prefix)
            available = max_length - len(prefix)
            if available <= 0:
                return prefix[:max_length]
            base = base[:available].rstrip("-_") or "item"
            return f"{prefix}{base}"

        return PathHelper.sanitize_component(base[:max_length])

    @staticmethod
    def course_dir_name(course_url: str, course_title: str) -> str:
        """Prefer the short course code from the URL for the course directory."""
        parsed = urlparse(course_url)
        course_id = parsed.path.rstrip("/").split("/")[-1].lower()
        course_id = PathHelper.sanitize_component(course_id)
        if course_id:
            return course_id[:MAX_COURSE_DIR_LENGTH]
        return PathHelper.shorten_title(
            course_title, fallback=course_title, max_length=MAX_COURSE_DIR_LENGTH
        )

    @staticmethod
    def ensure_path_length(path: str, max_length: int = MAX_FULL_PATH_LENGTH) -> str:
        """Trim the file stem if the full path would become too long."""
        normalized_path = os.path.normpath(path)
        if len(normalized_path) <= max_length:
            return path

        directory, filename = os.path.split(normalized_path)
        stem, extension = os.path.splitext(filename)
        overflow = len(normalized_path) - max_length
        shortened_stem_length = len(stem) - overflow

        if shortened_stem_length < 1:
            shortened_stem = "file"
        else:
            shortened_stem_length = max(8, shortened_stem_length)
            shortened_stem = stem[:shortened_stem_length].rstrip("-_. ") or "file"

        return os.path.join(directory, f"{shortened_stem}{extension}")


# =============================================================================
# HTTP Client
# =============================================================================


class HttpClient:
    """HTTP client for making requests with consistent configuration."""

    def __init__(self, headers: Optional[dict] = None, timeout: int = REQUEST_TIMEOUT):
        self.headers = headers or DEFAULT_HEADERS
        self.timeout = timeout

    def get(self, url: str, **kwargs) -> requests.Response:
        """Make a GET request with default headers and timeout."""
        request_headers = kwargs.pop("headers", self.headers)
        request_timeout = kwargs.pop("timeout", self.timeout)
        return requests.get(
            url, headers=request_headers, timeout=request_timeout, **kwargs
        )


# =============================================================================
# Catalog Service
# =============================================================================


class CatalogService:
    """Service for interacting with the Microsoft Learn Catalog API."""

    def __init__(self, http_client: Optional[HttpClient] = None):
        self.http_client = http_client or HttpClient(timeout=CATALOG_TIMEOUT)
        self._catalog: Optional[dict] = None

    def fetch(self) -> Optional[dict]:
        """Fetch the Microsoft Learn Catalog API."""
        try:
            print("Fetching Microsoft Learn catalog (this may take a moment)...")
            response = self.http_client.get(CATALOG_API_URL)
            response.raise_for_status()
            self._catalog = response.json()
            return self._catalog
        except requests.RequestException as e:
            print(f"Error fetching catalog: {e}")
            return None

    @property
    def catalog(self) -> Optional[dict]:
        """Get the cached catalog, fetching if necessary."""
        if self._catalog is None:
            self._catalog = self.fetch()
        return self._catalog

    def get_course_learning_paths(self, course_url: str) -> list[str]:
        """Extract all learning path URLs for a given course."""
        catalog = self.catalog
        if not catalog:
            return []

        course_id = self._extract_id_from_url(course_url)
        target_course = self._find_course_by_id(course_id, catalog.get("courses", []))

        if not target_course:
            print(f"Course '{course_id}' not found in catalog.")
            return []

        study_guide = target_course.get("study_guide", [])
        lp_uids = [
            item["uid"] for item in study_guide if item.get("type") == "learningPath"
        ]
        lp_lookup = {
            lp.get("uid"): lp.get("url") for lp in catalog.get("learningPaths", [])
        }
        return [
            self._clean_url(lp_lookup[uid])
            for uid in lp_uids
            if lp_lookup.get(uid)
        ]

    def get_learning_path_modules(self, path_url: str) -> list[str]:
        """Get all module URLs for a learning path."""
        catalog = self.catalog
        if not catalog:
            return []

        path_name = self._extract_id_from_url(path_url)
        target_lp = self._find_learning_path_by_name(
            path_name, catalog.get("learningPaths", [])
        )
        if not target_lp:
            return []

        module_uids = target_lp.get("modules", [])
        module_lookup = {
            mod.get("uid"): mod.get("url") for mod in catalog.get("modules", [])
        }
        return [
            self._clean_url(module_lookup[uid])
            for uid in module_uids
            if module_lookup.get(uid)
        ]

    @staticmethod
    def _extract_id_from_url(url: str) -> str:
        """Extract the last path segment from a URL, ignoring query strings."""
        return urlparse(url).path.rstrip("/").split("/")[-1]

    @staticmethod
    def _clean_url(url: str) -> str:
        """Remove query parameters from URL."""
        parsed = urlparse(url)
        return f"{parsed.scheme}://{parsed.netloc}{parsed.path}"

    @staticmethod
    def _find_course_by_id(course_id: str, courses: list[dict]) -> Optional[dict]:
        """Find a course by its ID using exact or suffix matching.

        Avoids false positives from substring matching (e.g. 'az-10' matching 'az-104t00').
        """
        course_id_lower = course_id.lower()
        # Pass 1: exact match or exact last dotted segment
        for course in courses:
            uid = course.get("uid", "").lower()
            last_segment = uid.rsplit(".", 1)[-1] if "." in uid else uid
            if uid == course_id_lower or last_segment == course_id_lower:
                return course
        # Pass 2: uid ends with the course id (e.g. 'learn.microsoft.com.az-140t00')
        for course in courses:
            uid = course.get("uid", "").lower()
            if uid.endswith(course_id_lower):
                return course
        return None

    @staticmethod
    def _find_learning_path_by_name(
        path_name: str, learning_paths: list[dict]
    ) -> Optional[dict]:
        """Find a learning path by exact last URL segment match."""
        path_name_lower = path_name.lower()
        for lp in learning_paths:
            lp_url = lp.get("url", "")
            last_segment = urlparse(lp_url).path.rstrip("/").split("/")[-1].lower()
            if last_segment == path_name_lower:
                return lp
        return None


# =============================================================================
# Content Service
# =============================================================================


class ContentService:
    """Service for fetching and processing web page content."""

    def __init__(self, http_client: Optional[HttpClient] = None):
        self.http_client = http_client or HttpClient()

    def fetch_page(self, url: str) -> PageContent:
        """Fetch and extract main content from a web page."""
        try:
            response = self.http_client.get(url)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "html.parser")
            title = self._extract_title(soup)
            content = self._extract_content(soup, url)
            return PageContent(title=title, content=content, url=url)
        except requests.RequestException as e:
            print(f"      Warning: Could not fetch {url}: {e}")
            return PageContent(
                title="Error",
                content=f"<p>Error loading content: {html.escape(str(e))}</p>",
                url=url,
            )

    def fetch_unit_links(self, module_url: str) -> list[str]:
        """Fetch all unit links from a module page."""
        try:
            response = self.http_client.get(module_url)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "html.parser")
            matching_links: set[str] = set()

            # Ensure the prefix ends with "/" so we don't accidentally match
            # other modules whose URL starts with the same characters.
            module_prefix = module_url.rstrip("/") + "/"

            for link in soup.find_all("a", href=True):
                full_url = urljoin(module_url, link["href"])
                parsed = urlparse(full_url)
                clean_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
                if clean_url.startswith(module_prefix):
                    matching_links.add(clean_url.rstrip("/"))

            return sorted(matching_links, key=self._extract_sort_key)
        except requests.RequestException as e:
            print(f"      Warning: Could not fetch units from {module_url}: {e}")
            return []

    @staticmethod
    def _extract_title(soup: BeautifulSoup) -> str:
        title_tag = soup.find("h1")
        return title_tag.get_text(strip=True) if title_tag else "Untitled"

    @staticmethod
    def _extract_content(soup: BeautifulSoup, base_url: str) -> str:
        content_div = (
            soup.find("article")
            or soup.find("main")
            or soup.find("div", class_="content")
        )
        if not content_div:
            return f"<p>Could not extract content from {html.escape(base_url)}</p>"
        ContentService._clean_navigation_elements(content_div)
        ContentService._fix_image_urls(content_div, base_url)
        return str(content_div)

    @staticmethod
    def _clean_navigation_elements(content_div) -> None:
        for nav in content_div.find_all(["nav", "aside", "footer"]):
            nav.decompose()
        for selector in [
            ".font-size-sm.margin-top-md.display-none-print",
            ".button.button-clear.button-primary.button-sm.inner-focus",
        ]:
            for elem in content_div.select(selector):
                elem.decompose()
        for elem in content_div.find_all(
            class_=lambda x: x and "background-color-body" in x
        ):
            elem.decompose()

    @staticmethod
    def _fix_image_urls(content_div, base_url: str) -> None:
        for img in content_div.find_all("img"):
            src = img.get("src")
            if src:
                img["src"] = urljoin(base_url, src)
            srcset = img.get("srcset")
            if srcset:
                new_srcset = []
                for item in srcset.split(","):
                    parts = item.strip().split(" ")
                    if parts:
                        parts[0] = urljoin(base_url, parts[0])
                        new_srcset.append(" ".join(parts))
                img["srcset"] = ", ".join(new_srcset)

    @staticmethod
    def _extract_sort_key(url: str) -> int | float:
        match = re.search(r"/([^/]+)$", url)
        if match:
            num_match = re.match(r"^(\d+)", match.group(1))
            if num_match:
                return int(num_match.group(1))
        return float("inf")


# =============================================================================
# HTML Generator
# =============================================================================


class HtmlGenerator:
    """Generator for creating combined HTML documents."""

    def __init__(self, content_service: Optional[ContentService] = None):
        self.content_service = content_service or ContentService()

    def generate_module_html(
        self,
        module_url: str,
        unit_links: list[str],
        output_dir: str,
        numbered_prefix: str,
    ) -> str:
        """Generate a combined HTML file with all unit contents."""
        module_data = self.content_service.fetch_page(module_url)
        module_slug_fallback = urlparse(module_url).path.rstrip("/").split("/")[-1]
        file_stem = PathHelper.shorten_title(
            module_data.title,
            fallback=module_slug_fallback,
            max_length=MAX_MODULE_FILE_STEM_LENGTH,
            prefix=f"{numbered_prefix}-",
        )
        output_file = PathHelper.ensure_path_length(
            os.path.join(output_dir, f"{file_stem}.html")
        )
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(self._build_html(module_data, unit_links))
        return output_file

    def _build_html(self, module_data: PageContent, unit_links: list[str]) -> str:
        sections = []
        for i, link in enumerate(unit_links, 1):
            page_data = self.content_service.fetch_page(link)
            sections.append(self._build_section(i, page_data))
        return self._build_document(module_data.title, sections)

    def _build_section(self, index: int, page_data: PageContent) -> str:
        safe_title = html.escape(page_data.title)
        safe_url = html.escape(page_data.url)
        return f"""
    <div class="section">
        <div class="section-header">
            <h2>{index}. {safe_title}</h2>
            <a href="{safe_url}">{safe_url}</a>
        </div>
        <div class="content">{page_data.content}</div>
    </div>"""

    def _build_document(self, title: str, sections: list[str]) -> str:
        safe_title = html.escape(title)
        sections_html = "\n".join(sections)
        return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{safe_title}</title>
    <style>{HTML_STYLES}</style>
</head>
<body>
    <h1>{safe_title}</h1>
{sections_html}
</body>
</html>"""


# =============================================================================
# PDF Generator
# =============================================================================


class PdfGenerator:
    """Generator for converting HTML to PDF using Playwright."""

    @staticmethod
    async def _convert_html_to_pdf(html_file: str, pdf_file: str) -> str:
        async with async_playwright() as p:
            browser = await p.chromium.launch()
            page = await browser.new_page()
            # Path.as_uri() produces a correctly encoded file:/// URI on all platforms,
            # including Windows paths with backslashes and spaces.
            await page.goto(Path(os.path.abspath(html_file)).as_uri())
            # "load" is reliable for local HTML; "networkidle" can hang on
            # slow or blocked external resources (e.g. images from learn.microsoft.com).
            await page.wait_for_load_state("load")
            await page.pdf(
                path=pdf_file,
                format="A4",
                margin={"top": "20px", "right": "20px", "bottom": "20px", "left": "20px"},
                print_background=True,
            )
            await browser.close()
        return pdf_file

    def generate(self, html_file: str) -> Optional[str]:
        """Generate a PDF from an HTML file, handling errors gracefully."""
        stem, _ = os.path.splitext(html_file)
        pdf_file = PathHelper.ensure_path_length(f"{stem}.pdf")
        try:
            asyncio.run(self._convert_html_to_pdf(html_file, pdf_file))
            return pdf_file
        except Exception as e:
            print(f"      Warning: PDF generation failed: {e}")
            return None


# =============================================================================
# Index Generator
# =============================================================================


class IndexGenerator:
    """Generates an index.html overview page in the course root directory."""

    @staticmethod
    def generate(
        course_title: str,
        course_url: str,
        course_output_dir: str,
        lp_entries: list[LearningPathIndexEntry],
    ) -> str:
        """Write index.html to the course root and return its path."""
        index_path = os.path.join(course_output_dir, "index.html")
        content = IndexGenerator._build(
            course_title, course_url, course_output_dir, lp_entries
        )
        with open(index_path, "w", encoding="utf-8") as f:
            f.write(content)
        return index_path

    @staticmethod
    def _build(
        course_title: str,
        course_url: str,
        course_output_dir: str,
        lp_entries: list[LearningPathIndexEntry],
    ) -> str:
        safe_title = html.escape(course_title)
        safe_url = html.escape(course_url)
        lp_blocks = "\n".join(
            IndexGenerator._lp_block(lp, course_output_dir)
            for lp in lp_entries
        )
        return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{safe_title}</title>
    <style>{INDEX_STYLES}</style>
</head>
<body>
    <h1>{safe_title}</h1>
    <p class="course-url">
        Source: <a href="{safe_url}" target="_blank">{safe_url}</a>
    </p>
{lp_blocks}
    <p class="footer">Generated by Microsoft Learn Course Extractor</p>
</body>
</html>"""

    @staticmethod
    def _lp_block(lp: LearningPathIndexEntry, course_output_dir: str) -> str:
        safe_title = html.escape(lp.lp_title)
        safe_url = html.escape(lp.lp_url)
        module_items = "\n".join(
            IndexGenerator._module_item(mod, course_output_dir)
            for mod in lp.modules
        )
        return f"""    <div class="lp-block">
        <div class="lp-header">
            <span class="lp-num">{lp.lp_index:02d}</span>
            <a href="{safe_url}" target="_blank">{safe_title}</a>
        </div>
        <ul class="module-list">
{module_items}
        </ul>
    </div>"""

    @staticmethod
    def _module_item(mod: ModuleIndexEntry, course_output_dir: str) -> str:
        safe_title = html.escape(mod.module_title)
        num_label = f"{mod.lp_index:02d}-{mod.module_index:02d}"
        links_html = ""

        if mod.html_file and os.path.exists(mod.html_file):
            rel = os.path.relpath(mod.html_file, course_output_dir).replace("\\", "/")
            links_html += f'<a class="link-html" href="{html.escape(rel)}">HTML</a>'

        if mod.pdf_file and os.path.exists(mod.pdf_file):
            rel = os.path.relpath(mod.pdf_file, course_output_dir).replace("\\", "/")
            links_html += f'<a class="link-pdf" href="{html.escape(rel)}">PDF</a>'

        return f"""            <li>
                <span class="module-num">{num_label}</span>
                <span class="module-title">{safe_title}</span>
                <span class="file-links">{links_html}</span>
            </li>"""


# =============================================================================
# Course Processor
# =============================================================================


class CourseProcessor:
    """Main processor for extracting and generating course content."""

    def __init__(
        self,
        output_format: str = OUTPUT_FORMAT_BOTH,
        catalog_service: Optional[CatalogService] = None,
        content_service: Optional[ContentService] = None,
        html_generator: Optional[HtmlGenerator] = None,
        pdf_generator: Optional[PdfGenerator] = None,
    ):
        self.output_format = output_format
        self.catalog_service = catalog_service or CatalogService()
        self.content_service = content_service or ContentService()
        self.html_generator = html_generator or HtmlGenerator(self.content_service)
        self.pdf_generator = pdf_generator or PdfGenerator()

    def process_course(
        self, course_url: str, output_base: str = DEFAULT_OUTPUT_BASE_DIR
    ) -> list[str]:
        """Process a course and generate all output files."""
        print(f"\nFetching learning paths from: {course_url}")
        print("=" * 80)

        # Resolve to absolute path so all console output shows full paths.
        output_base = os.path.abspath(output_base)
        course_title = self._fetch_course_title(course_url)
        course_dir_name = PathHelper.course_dir_name(course_url, course_title)
        course_output_dir = os.path.join(output_base, course_dir_name)

        catalog = self.catalog_service.fetch()
        if not catalog:
            print("Error: Failed to fetch the Microsoft Learn catalog. Check your internet connection.")
            return []

        paths = self.catalog_service.get_course_learning_paths(course_url)
        if not paths:
            print("\nNo learning paths found for this course.")
            print("Make sure the URL points to a valid Microsoft Learn course page.")
            return []

        self._display_learning_paths(paths)

        os.makedirs(course_output_dir, exist_ok=True)
        print(f"\n  Course:           {course_title}")
        print(f"  Course directory: {course_dir_name}")
        print(f"  Output directory: {course_output_dir}")
        print(f"  Output format:    {self.output_format.upper()}")

        lp_entries: list[LearningPathIndexEntry] = []
        for i, path_url in enumerate(paths, 1):
            lp_entry = self._process_learning_path(path_url, i, course_output_dir)
            lp_entries.append(lp_entry)

        # Write the course-level index.html
        index_file = IndexGenerator.generate(
            course_title, course_url, course_output_dir, lp_entries
        )

        print(f"\n{'=' * 80}")
        print("Done! Output is saved in:")
        print(f"  {course_output_dir}")
        print(f"  Index: {index_file}")

        return paths

    def _fetch_course_title(self, course_url: str) -> str:
        try:
            return self.content_service.fetch_page(course_url).title
        except Exception as e:
            print(f"Warning: Could not fetch course title: {e}")
            return urlparse(course_url).path.rstrip("/").split("/")[-1]

    def _display_learning_paths(self, paths: list[str]) -> None:
        print(f"\nFound {len(paths)} learning path(s):\n")
        for i, path in enumerate(paths, 1):
            print(f"  {i}. {path}")
        print("\n" + "=" * 80)
        print("Creating directories and generating content...")

    def _process_learning_path(
        self, path_url: str, lp_index: int, output_base: str
    ) -> LearningPathIndexEntry:
        """Process a single learning path and return its index entry."""
        print(f"\nFetching learning path page: {path_url}")

        path_data = self.content_service.fetch_page(path_url)
        path_slug_fallback = urlparse(path_url).path.rstrip("/").split("/")[-1]
        numbered_name = PathHelper.shorten_title(
            path_data.title,
            fallback=path_slug_fallback,
            max_length=MAX_LEARNING_PATH_DIR_LENGTH,
            prefix=f"{lp_index:02d}-",
        )
        path_dir = os.path.join(output_base, numbered_name)
        os.makedirs(path_dir, exist_ok=True)

        print(f"\nLearning Path: {path_data.title}")
        print(f"  Directory: {numbered_name}")
        print(f"  Created:   {path_dir}")

        lp_entry = LearningPathIndexEntry(
            lp_index=lp_index,
            lp_title=path_data.title,
            lp_url=path_url,
            lp_dir=path_dir,
        )

        modules = self.catalog_service.get_learning_path_modules(path_url)
        if modules:
            for j, module_url in enumerate(modules, 1):
                mod_entry = self._process_module(module_url, lp_index, j, path_dir)
                lp_entry.modules.append(mod_entry)
        else:
            print("    (No modules found for this learning path)")

        return lp_entry

    def _process_module(
        self, module_url: str, lp_index: int, module_index: int, path_dir: str
    ) -> ModuleIndexEntry:
        """Process a single module and return its index entry."""
        module_name = urlparse(module_url).path.rstrip("/").split("/")[-1]
        print(f"\n    Module: {module_name}")
        print("      Fetching units...")

        unit_links = self.content_service.fetch_unit_links(module_url)
        if not unit_links:
            print("      No units found for this module.")
            return ModuleIndexEntry(
                lp_index=lp_index,
                module_index=module_index,
                module_title=module_name,
                module_url=module_url,
                html_file=None,
                pdf_file=None,
            )

        print(f"      Found {len(unit_links)} unit(s)")

        # Filename prefix includes both LP number and module number: e.g. "06-02"
        numbered_prefix = f"{lp_index:02d}-{module_index:02d}"

        html_file: Optional[str] = self.html_generator.generate_module_html(
            module_url, unit_links, path_dir, numbered_prefix
        )
        pdf_file: Optional[str] = None

        if self.output_format in (OUTPUT_FORMAT_PDF, OUTPUT_FORMAT_BOTH):
            pdf_file = self.pdf_generator.generate(html_file)

        if self.output_format == OUTPUT_FORMAT_PDF:
            # PDF-only: remove the intermediate HTML file after conversion.
            if html_file and os.path.exists(html_file):
                try:
                    os.remove(html_file)
                except OSError as e:
                    print(f"      Warning: Could not remove intermediate HTML: {e}")
            html_file = None
            if pdf_file:
                print(f"      PDF:  {pdf_file}")
            else:
                print("      PDF generation failed; no output for this module.")
        elif self.output_format == OUTPUT_FORMAT_HTML:
            print(f"      HTML: {html_file}")
        else:  # both
            print(f"      HTML: {html_file}")
            if pdf_file:
                print(f"      PDF:  {pdf_file}")

        # Fetch the module title for the index page.
        module_title = module_name
        try:
            module_title = self.content_service.fetch_page(module_url).title
        except Exception:
            pass

        return ModuleIndexEntry(
            lp_index=lp_index,
            module_index=module_index,
            module_title=module_title,
            module_url=module_url,
            html_file=html_file,
            pdf_file=pdf_file,
        )


# =============================================================================
# User Input
# =============================================================================


def validate_course_url(url: str) -> bool:
    """Return True if the URL looks like a valid Microsoft Learn course URL."""
    try:
        parsed = urlparse(url)
        return (
            parsed.scheme in ("http", "https")
            and "learn.microsoft.com" in parsed.netloc
            and LEARN_COURSE_PATH_PREFIX in parsed.path
        )
    except Exception:
        return False


def get_course_url_from_user() -> str:
    """Prompt the user for a Microsoft Learn course URL."""
    print(
        "Enter the Microsoft Learn course URL."
        "\n  Example: https://learn.microsoft.com/en-us/training/courses/az-140t00"
    )
    while True:
        course_url = input("> ").strip().strip('"')
        if not course_url:
            print("  Please enter a URL.")
            continue
        if not validate_course_url(course_url):
            print(
                "  That does not look like a valid Microsoft Learn course URL.\n"
                "  Expected format: https://learn.microsoft.com/.../training/courses/<course-code>\n"
                "  Please try again."
            )
            continue
        return course_url


def get_output_base_from_user() -> str:
    """Prompt the user for the output base directory."""
    default_display = os.path.abspath(DEFAULT_OUTPUT_BASE_DIR)
    print(
        f"\nEnter the output directory"
        f" (press Enter to use current directory: {default_display}):"
    )
    output_base = input("> ").strip().strip('"')

    if not output_base:
        return DEFAULT_OUTPUT_BASE_DIR

    if not os.path.exists(output_base):
        try:
            os.makedirs(output_base, exist_ok=True)
            print(f"  Created output directory: {os.path.abspath(output_base)}")
        except OSError as e:
            print(f"  Warning: Could not create directory '{output_base}': {e}")
            print(f"  Falling back to current directory: {default_display}")
            return DEFAULT_OUTPUT_BASE_DIR

    return output_base


def get_output_format_from_user() -> str:
    """Ask the user which output format to generate."""
    print(
        "\nWhich output format do you want?"
        "\n  1 = HTML only"
        "\n  2 = PDF only  (HTML is generated as intermediate and then removed)"
        "\n  3 = Both HTML and PDF  [default]"
    )
    mapping = {
        "1": OUTPUT_FORMAT_HTML,
        "2": OUTPUT_FORMAT_PDF,
        "3": OUTPUT_FORMAT_BOTH,
        "":  OUTPUT_FORMAT_BOTH,
    }
    labels = {
        OUTPUT_FORMAT_HTML: "HTML only",
        OUTPUT_FORMAT_PDF:  "PDF only",
        OUTPUT_FORMAT_BOTH: "HTML + PDF",
    }
    while True:
        choice = input("> ").strip()
        if choice in mapping:
            selected = mapping[choice]
            print(f"  Output format: {labels[selected]}")
            return selected
        print("  Please enter 1, 2 or 3 (or press Enter for the default).")


# =============================================================================
# Main Entry Point
# =============================================================================


def main() -> list[str]:
    """Main entry point."""
    course_url = get_course_url_from_user()
    output_base = get_output_base_from_user()
    output_format = get_output_format_from_user()

    processor = CourseProcessor(output_format=output_format)
    return processor.process_course(course_url, output_base)


if __name__ == "__main__":
    main()
