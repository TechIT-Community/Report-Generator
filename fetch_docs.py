import requests
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
import time

BASE_URL = "https://customtkinter.tomschimansky.com/"
visited = set()
OUTPUT_FILE = "customtkinter_docs.txt"


def extract_readable_text(soup):
    """
    Extract readable documentation-style text while:
    - preserving paragraphs
    - preserving code blocks EXACTLY
    - avoiding newline spam
    """
    parts = []
    main = soup.body or soup

    for elem in main.find_all(
        ["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "pre"],
        recursive=True,
    ):
        # ---- CODE BLOCKS (PRESERVE EXACTLY) ----
        if elem.name == "pre":
            code = elem.text.rstrip("\n")
            if code.strip():
                parts.append("\n```")
                parts.append(code)
                parts.append("```")
            continue

        # ---- NORMAL TEXT BLOCKS ----
        text = " ".join(elem.stripped_strings)
        if text:
            parts.append(text)

    return "\n\n".join(parts)


def crawl(url):
    if url in visited:
        return
    visited.add(url)

    print(f"Crawling: {url}")

    try:
        res = requests.get(url, timeout=10)
        res.raise_for_status()
    except Exception as e:
        print(f"Failed to fetch {url}: {e}")
        return

    soup = BeautifulSoup(res.text, "html.parser")

    content = extract_readable_text(soup)

    with open(OUTPUT_FILE, "a", encoding="utf-8") as f:
        f.write(f"\n\n===== {url} =====\n\n")
        f.write(content)

    # ---- FOLLOW INTERNAL LINKS ----
    for a in soup.find_all("a", href=True):
        abs_link = urljoin(url, a["href"])
        parsed = urlparse(abs_link)

        if parsed.netloc == urlparse(BASE_URL).netloc:
            clean = parsed.scheme + "://" + parsed.netloc + parsed.path
            if clean not in visited:
                crawl(clean)

    time.sleep(0.3)  # be polite


if __name__ == "__main__":
    open(OUTPUT_FILE, "w", encoding="utf-8").close()
    crawl(BASE_URL)
    print("Done crawling.")
