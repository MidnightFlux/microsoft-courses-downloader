"""
Script to extract learning path links from a Microsoft Learn course page.
Uses the Microsoft Learn Catalog API.

This version shortens directory and file names so it behaves better on Windows
systems that still hit MAX_PATH limitations.
"""

import os
import re
import requests
import asyncio
from dataclasses import dataclass
from typing import Optional
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from playwright.async_api import async_playwright


# =============================================================================
# Constants
# =============================================================================

DEFAULT_COURSE_URL = "https://learn.microsoft.com/en-us/training/courses/ai-102t00"
CATALOG_API_URL = "https://learn.microsoft.com/api/catalog/"
OUTPUT_BASE_DIR = "output"
REQUEST_TIMEOUT = 30
CATALOG_TIMEOUT = 60

MAX_COURSE_DIR_LENGTH = 32
MAX_LEARNING_PATH_DIR_LENGTH = 48
MAX_MODULE_FILE_STEM_LENGTH = 64
MAX_FULL_PATH_LENGTH = 235

DEFAULT_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.0"
}

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

COMMON_PHRASE_REPLACEMENTS = {
    "microsoft dynamics 365 customer service": "",
    "dynamics 365 customer service": "",
    "microsoft dynamics 365": "",
    "microsoft learn": "",
    "microsoft": "",
    "customer service": "",
    "service-level": "service level",
    "service level": "sla",
    "agreements": "agreements",
}

STOPWORDS = {
    "a",
    "an",
    "and",
    "by",
    "for",
    "from",
    "in",
    "into",
    "of",
    "on",
    "or",
    "the",
    "to",
    "using",
    "with",
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
        words = [word for word in text.split() if word and word not in STOPWORDS]
        slug = "-".join(words)
        slug = re.sub(r"-+", "-", slug).strip("-")
        return slug

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
        if len(base) > max_length:
            base = base[:max_length].rstrip("-_")

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
        course_id = course_url.rstrip("/").split("/")[-1].lower()
        course_id = PathHelper.sanitize_component(course_id)

        if course_id:
            return course_id[:MAX_COURSE_DIR_LENGTH]

        return PathHelper.shorten_title(
            course_title,
            fallback=course_title,
            max_length=MAX_COURSE_DIR_LENGTH,
        )

    @staticmethod
    def ensure_path_length(path: str, max_length: int = MAX_FULL_PATH_LENGTH) -> str:
        """Trim the file name if the full path would become too long."""
        normalized_path = os.path.normpath(path)
        if len(normalized_path) <= max_length:
            return path

        directory, filename = os.path.split(path)
        stem, extension = os.path.splitext(filename)
        overflow = len(normalized_path) - max_length
        shortened_stem_length = max(8, len(stem) - overflow)
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
        learning_path_uids = [
            item["uid"] for item in study_guide if item.get("type") == "learningPath"
        ]

        lp_lookup = {
            lp.get("uid"): lp.get("url") for lp in catalog.get("learningPaths", [])
        }

        path_urls = []
        for uid in learning_path_uids:
            url = lp_lookup.get(uid)
            if url:
                path_urls.append(self._clean_url(url))

        return path_urls

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

        module_urls = []
        for uid in module_uids:
            url = module_lookup.get(uid)
            if url:
                module_urls.append(self._clean_url(url))

        return module_urls

    @staticmethod
    def _extract_id_from_url(url: str) -> str:
        """Extract the ID from a URL."""
        return url.rstrip("/").split("/")[-1]

    @staticmethod
    def _clean_url(url: str) -> str:
        """Remove query parameters from URL."""
        return url.split("?")[0]

    @staticmethod
    def _find_course_by_id(course_id: str, courses: list[dict]) -> Optional[dict]:
        """Find a course by its ID (case-insensitive partial match)."""
        course_id_lower = course_id.lower()
        for course in courses:
            uid = course.get("uid", "")
            if course_id_lower in uid.lower():
                return course
        return None

    @staticmethod
    def _find_learning_path_by_name(
        path_name: str, learning_paths: list[dict]
    ) -> Optional[dict]:
        """Find a learning path by its name in the URL."""
        for lp in learning_paths:
            lp_url = lp.get("url", "")
            if path_name in lp_url:
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
            print(f"Error fetching {url}: {e}")
            return PageContent(
                title="Error", content=f"<p>Error loading content: {e}</p>", url=url
            )

    def fetch_unit_links(self, module_url: str) -> list[str]:
        """Fetch all unit links from a module URL."""
        try:
            response = self.http_client.get(module_url)
            response.raise_for_status()

            soup = BeautifulSoup(response.content, "html.parser")
            matching_links = set()

            for link in soup.find_all("a", href=True):
                href = link["href"]
                full_url = urljoin(module_url, href)
                parsed = urlparse(full_url)
                clean_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"

                if clean_url.startswith(module_url) and clean_url != module_url:
                    matching_links.add(clean_url)

            return sorted(matching_links, key=self._extract_sort_key)

        except requests.RequestException as e:
            print(f"Error fetching units from {module_url}: {e}")
            return []

    @staticmethod
    def _extract_title(soup: BeautifulSoup) -> str:
        """Extract the page title from the h1 tag."""
        title_tag = soup.find("h1")
        return title_tag.get_text(strip=True) if title_tag else "Untitled"

    @staticmethod
    def _extract_content(soup: BeautifulSoup, base_url: str) -> str:
        """Extract and clean the main content from the page."""
        content_div = (
            soup.find("article")
            or soup.find("main")
            or soup.find("div", class_="content")
        )

        if not content_div:
            return f"<p>Could not extract content from {base_url}</p>"

        ContentService._clean_navigation_elements(content_div)
        ContentService._fix_image_urls(content_div, base_url)

        return str(content_div)

    @staticmethod
    def _clean_navigation_elements(content_div) -> None:
        """Remove navigation and UI elements from content."""
        for nav in content_div.find_all(["nav", "aside", "footer"]):
            nav.decompose()

        selectors_to_remove = [
            ".font-size-sm.margin-top-md.display-none-print",
            ".button.button-clear.button-primary.button-sm.inner-focus",
        ]
        for selector in selectors_to_remove:
            for elem in content_div.select(selector):
                elem.decompose()

        for elem in content_div.find_all(
            class_=lambda x: x and "background-color-body" in x
        ):
            elem.decompose()

    @staticmethod
    def _fix_image_urls(content_div, base_url: str) -> None:
        """Convert relative image URLs to absolute."""
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
        """Extract numeric prefix from URL path for sorting."""
        match = re.search(r"/([^/]+)$", url)
        if match:
            segment = match.group(1)
            num_match = re.match(r"^(\d+)", segment)
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
        module_slug_fallback = module_url.rstrip("/").split("/")[-1]
        file_stem = PathHelper.shorten_title(
            module_data.title,
            fallback=module_slug_fallback,
            max_length=MAX_MODULE_FILE_STEM_LENGTH,
            prefix=f"{numbered_prefix}-",
        )

        html_filename = f"{file_stem}.html"
        output_file = PathHelper.ensure_path_length(
            os.path.join(output_dir, html_filename)
        )

        html_content = self._build_html(module_data, unit_links)

        with open(output_file, "w", encoding="utf-8") as f:
            f.write(html_content)

        return output_file

    def _build_html(self, module_data: PageContent, unit_links: list[str]) -> str:
        """Build the complete HTML document."""
        sections = []
        for i, link in enumerate(unit_links, 1):
            page_data = self.content_service.fetch_page(link)
            sections.append(self._build_section(i, page_data))

        return self._build_document(module_data.title, sections)

    def _build_section(self, index: int, page_data: PageContent) -> str:
        """Build a single section HTML."""
        return f"""
    <div class="section">
        <div class="section-header">
            <h2>{index}. {page_data.title}</h2>
            <a href="{page_data.url}">{page_data.url}</a>
        </div>
        <div class="content">{page_data.content}</div>
    </div>"""

    def _build_document(self, title: str, sections: list[str]) -> str:
        """Build the complete HTML document structure."""
        sections_html = "\n".join(sections)

        return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>{HTML_STYLES}</style>
</head>
<body>
    <h1>{title}</h1>
{sections_html}
</body>
</html>"""


# =============================================================================
# PDF Generator
# =============================================================================


class PdfGenerator:
    """Generator for converting HTML to PDF."""

    @staticmethod
    async def convert_html_to_pdf(html_file: str, pdf_file: str) -> str:
        """Convert an HTML file to PDF using Playwright."""
        async with async_playwright() as p:
            browser = await p.chromium.launch()
            page = await browser.new_page()

            html_path = os.path.abspath(html_file)
            await page.goto(f"file:///{html_path}")
            await page.wait_for_load_state("networkidle")

            await page.pdf(
                path=pdf_file,
                format="A4",
                margin={
                    "top": "20px",
                    "right": "20px",
                    "bottom": "20px",
                    "left": "20px",
                },
                print_background=True,
            )

            await browser.close()

        return pdf_file

    def generate(self, html_file: str) -> Optional[str]:
        """Generate PDF from HTML file, handling errors gracefully."""
        pdf_file = PathHelper.ensure_path_length(html_file.replace(".html", ".pdf"))
        try:
            asyncio.run(self.convert_html_to_pdf(html_file, pdf_file))
            return pdf_file
        except Exception as e:
            print(f"      Warning: Failed to generate PDF: {e}")
            return None


# =============================================================================
# Course Processor
# =============================================================================


class CourseProcessor:
    """Main processor for extracting and generating course content."""

    def __init__(
        self,
        catalog_service: Optional[CatalogService] = None,
        content_service: Optional[ContentService] = None,
        html_generator: Optional[HtmlGenerator] = None,
        pdf_generator: Optional[PdfGenerator] = None,
    ):
        self.catalog_service = catalog_service or CatalogService()
        self.content_service = content_service or ContentService()
        self.html_generator = html_generator or HtmlGenerator(self.content_service)
        self.pdf_generator = pdf_generator or PdfGenerator()

    def process_course(
        self, course_url: str, output_base: str = OUTPUT_BASE_DIR
    ) -> list[str]:
        """Process a course and generate all output files."""
        print(f"\nFetching learning paths from: {course_url}")
        print("=" * 80)

        course_title = self._fetch_course_title(course_url)
        course_dir_name = PathHelper.course_dir_name(course_url, course_title)
        course_output_dir = os.path.join(output_base, course_dir_name)

        catalog = self.catalog_service.fetch()
        if not catalog:
            print("Failed to fetch catalog.")
            return []

        paths = self.catalog_service.get_course_learning_paths(course_url)

        if not paths:
            print("\nNo learning paths found.")
            return []

        self._display_learning_paths(paths)

        os.makedirs(course_output_dir, exist_ok=True)
        print(f"\n  Course: {course_title}")
        print(f"  Course directory: {course_dir_name}")
        print(f"  Output directory: {course_output_dir}/")

        for i, path in enumerate(paths, 1):
            self._process_learning_path(path, i, course_output_dir)

        print(f"\n{'=' * 80}")
        print(f"All done! Output is in '{course_output_dir}/'")

        return paths

    def _fetch_course_title(self, course_url: str) -> str:
        """Fetch the course page and extract the h1 title."""
        try:
            page_content = self.content_service.fetch_page(course_url)
            return page_content.title
        except Exception as e:
            print(f"Warning: Could not fetch course title: {e}")
            return course_url.rstrip("/").split("/")[-1]

    def _display_learning_paths(self, paths: list[str]) -> None:
        """Display the found learning paths."""
        print(f"\nFound {len(paths)} learning path(s):\n")
        for i, path in enumerate(paths, 1):
            print(f"{i}. {path}")
        print("\n" + "=" * 80)
        print("Creating directories and generating content...")

    def _process_learning_path(
        self, path_url: str, index: int, output_base: str
    ) -> None:
        """Process a single learning path."""
        print(f"\nFetching learning path page: {path_url}")

        path_data = self.content_service.fetch_page(path_url)
        path_slug_fallback = path_url.rstrip("/").split("/")[-1]
        numbered_name = PathHelper.shorten_title(
            path_data.title,
            fallback=path_slug_fallback,
            max_length=MAX_LEARNING_PATH_DIR_LENGTH,
            prefix=f"{index:02d}-",
        )
        path_dir = os.path.join(output_base, numbered_name)

        os.makedirs(path_dir, exist_ok=True)
        print(f"\nLearning Path: {path_data.title}")
        print(f"  Directory: {numbered_name}")
        print(f"  Created: {path_dir}/")

        modules = self.catalog_service.get_learning_path_modules(path_url)

        if modules:
            for j, module_url in enumerate(modules, 1):
                self._process_module(module_url, j, path_dir)
        else:
            print("    (No modules found for this learning path)")

    def _process_module(self, module_url: str, index: int, path_dir: str) -> None:
        """Process a single module."""
        module_name = module_url.rstrip("/").split("/")[-1]
        print(f"\n    Module: {module_name}")

        print("      Fetching units...")
        unit_links = self.content_service.fetch_unit_links(module_url)

        if not unit_links:
            print("      No units found for this module")
            return

        print(f"      Found {len(unit_links)} unit(s)")

        numbered_prefix = f"{index:02d}"
        html_file = self.html_generator.generate_module_html(
            module_url, unit_links, path_dir, numbered_prefix
        )
        print(f"      Generated: {html_file}")

        pdf_file = self.pdf_generator.generate(html_file)
        if pdf_file:
            print(f"      Generated: {pdf_file}")


# =============================================================================
# User Input
# =============================================================================


def get_course_url_from_user(default_url: str = DEFAULT_COURSE_URL) -> str:
    """Prompt user for course URL with default value."""
    print(
        f"Enter the Microsoft Learn course URL (press Enter to use default: {default_url}):"
    )
    course_url = input("> ").strip()
    return course_url if course_url else default_url


# =============================================================================
# Main Entry Point
# =============================================================================


def main() -> list[str]:
    """Main entry point."""
    course_url = get_course_url_from_user()
    processor = CourseProcessor()
    return processor.process_course(course_url)


if __name__ == "__main__":
    main()
