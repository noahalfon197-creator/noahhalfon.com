#!/usr/bin/env python3
"""
Auto-update script for Noah's website.
Watches the newsletters folder and automatically converts new/changed files.
Also removes newsletters when files are deleted.
"""

import mammoth
import os
import re
import base64
import time
import urllib.request
import ssl
import xml.etree.ElementTree as ET
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# SSL context for macOS certificate issues
SSL_CONTEXT = ssl.create_default_context()
SSL_CONTEXT.check_hostname = False
SSL_CONTEXT.verify_mode = ssl.CERT_NONE

# Blogger RSS feed URL
BLOGGER_RSS = "https://apollo1star.blogspot.com/feeds/posts/default?alt=rss"

# Paths
NEWSLETTERS_INPUT = os.path.expanduser("~/Desktop/newsletters")
NEWSLETTERS_OUTPUT = os.path.expanduser("~/Desktop/website/newsletters")
THUMB_DIR = os.path.expanduser("~/Desktop/website/images/newsletters")
INDEX_FILE = os.path.expanduser("~/Desktop/website/index.html")

os.makedirs(NEWSLETTERS_OUTPUT, exist_ok=True)
os.makedirs(THUMB_DIR, exist_ok=True)

def fetch_blogger_posts():
    """Fetch posts from blogger RSS feed and return a dict mapping date ranges to titles."""
    blogger_posts = {}
    try:
        with urllib.request.urlopen(BLOGGER_RSS, timeout=10, context=SSL_CONTEXT) as response:
            rss_content = response.read()

        root = ET.fromstring(rss_content)

        for item in root.findall('.//item'):
            title = item.find('title')
            pub_date = item.find('pubDate')

            if title is not None and pub_date is not None:
                title_text = title.text
                # Parse date like "Sun, 22 Feb 2026 00:00:00 +0000"
                try:
                    date_str = pub_date.text
                    parsed_date = datetime.strptime(date_str[:16], "%a, %d %b %Y")
                    date_formatted = parsed_date.strftime("%b %d, %Y")

                    # Extract date range from title (e.g., "Market Summary 2/16-2/20")
                    date_match = re.search(r'(\d{1,2}/\d{1,2})\s*-?\s*(\d{1,2}/\d{1,2})', title_text)
                    if date_match:
                        date_range_key = f"{date_match.group(1)}-{date_match.group(2)}"
                        # Normalize the key (remove spaces)
                        date_range_key = date_range_key.replace(' ', '')
                        blogger_posts[date_range_key] = {
                            'title': title_text,
                            'date': date_formatted
                        }
                except Exception as e:
                    print(f"  Warning: Could not parse blogger date: {e}")

        print(f"  Fetched {len(blogger_posts)} posts from blogger")
    except Exception as e:
        print(f"  Warning: Could not fetch blogger RSS: {e}")

    return blogger_posts

def extract_date_range(filename):
    """Extract date range from filename for matching with blogger."""
    # Match patterns like "2-16---2-20" or "2-16-2-20" or "2/16-2/20"
    match = re.search(r'(\d{1,2})[-/](\d{1,2})[-/\s]*[-/]?[-/\s]*(\d{1,2})[-/](\d{1,2})', filename)
    if match:
        return f"{match.group(1)}/{match.group(2)}-{match.group(3)}/{match.group(4)}"
    return None

def get_safename(filename):
    """Convert filename to safe HTML filename."""
    safename = filename.replace('.docx', '').replace(' ', '-').replace('_', '-')
    return re.sub(r'[^a-zA-Z0-9\-]', '', safename)

def convert_newsletter(filepath):
    """Convert a single newsletter docx to HTML."""
    filename = os.path.basename(filepath)

    if not filename.endswith('.docx') or filename.startswith('~') or 'TEMPLATE' in filename or 'PRICES' in filename:
        return None

    safename = get_safename(filename)

    try:
        with open(filepath, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value

        # Extract first image for thumbnail
        thumb_file = None
        img_match = re.search(r'<img[^>]+src="data:image/([^;]+);base64,([^"]+)"', html)
        if img_match:
            img_type = img_match.group(1)
            img_data = img_match.group(2)
            thumb_file = f"{safename}.{img_type}"
            thumb_path = os.path.join(THUMB_DIR, thumb_file)
            with open(thumb_path, 'wb') as tf:
                tf.write(base64.b64decode(img_data))

        # Extract colors from docx
        try:
            from docx import Document
            doc = Document(filepath)
            color_map = {}
            for para in doc.paragraphs:
                for run in para.runs:
                    if run.font.color and run.font.color.rgb:
                        rgb = run.font.color.rgb
                        hex_color = f"#{rgb}"
                        if run.text.strip():
                            color_map[run.text.strip()] = hex_color

            for text, color in color_map.items():
                if len(text) > 2:
                    escaped_text = re.escape(text)
                    html = re.sub(f'({escaped_text})', f'<span style="color:{color}">\\1</span>', html, count=1)
        except:
            pass

        full_html = f'''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{filename.replace('.docx', '')}</title>
    <style>
        body {{
            background: #fff;
            color: #000;
            font-family: 'Inter', -apple-system, sans-serif;
            max-width: 900px;
            margin: 0 auto;
            padding: 2rem;
            line-height: 1.7;
        }}
        img {{ max-width: 100%; height: auto; margin: 1rem 0; border-radius: 8px; }}
        table {{ border-collapse: collapse; width: 100%; margin: 1rem 0; }}
        th, td {{ border: 1px solid rgba(0,0,0,0.2); padding: 0.5rem; text-align: left; }}
        .back-link {{
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            margin-bottom: 2rem;
            color: #666;
            text-decoration: none;
            font-size: 0.9rem;
        }}
        .back-link:hover {{ color: #000; }}
        strong, b {{ font-weight: 600; }}
        h1, h2, h3 {{ margin: 1.5rem 0 1rem; }}
    </style>
</head>
<body>
    <a href="../index.html" class="back-link">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M19 12H5M12 19l-7-7 7-7"/>
        </svg>
        Back to Home
    </a>
    {html}
</body>
</html>'''

        output_path = os.path.join(NEWSLETTERS_OUTPUT, f"{safename}.html")
        with open(output_path, 'w') as out:
            out.write(full_html)

        mod_time = os.path.getmtime(filepath)

        return {
            'file': f"{safename}.html",
            'thumb': thumb_file,
            'title': filename.replace('.docx', ''),
            'mod_time': mod_time
        }

    except Exception as e:
        print(f"Error converting {filename}: {e}")
        return None

def delete_newsletter(filename):
    """Delete a newsletter's HTML and thumbnail."""
    safename = get_safename(filename)

    # Delete HTML file
    html_path = os.path.join(NEWSLETTERS_OUTPUT, f"{safename}.html")
    if os.path.exists(html_path):
        os.remove(html_path)
        print(f"  Deleted: {safename}.html")

    # Delete thumbnail (try common extensions)
    for ext in ['png', 'jpeg', 'jpg', 'gif']:
        thumb_path = os.path.join(THUMB_DIR, f"{safename}.{ext}")
        if os.path.exists(thumb_path):
            os.remove(thumb_path)
            print(f"  Deleted thumbnail: {safename}.{ext}")

def rebuild_all():
    """Rebuild all newsletters and update index."""
    print("Rebuilding all newsletters...")

    # Fetch blogger posts for title matching
    blogger_posts = fetch_blogger_posts()

    # Get list of current docx files
    current_files = set()
    for filename in os.listdir(NEWSLETTERS_INPUT):
        if filename.endswith('.docx') and not filename.startswith('~'):
            current_files.add(get_safename(filename))

    # Clean up old HTML files that no longer have source docx
    for html_file in os.listdir(NEWSLETTERS_OUTPUT):
        if html_file.endswith('.html'):
            safename = html_file.replace('.html', '')
            if safename not in current_files:
                os.remove(os.path.join(NEWSLETTERS_OUTPUT, safename + '.html'))
                print(f"  Removed orphan: {html_file}")
                # Also remove thumbnail
                for ext in ['png', 'jpeg', 'jpg', 'gif']:
                    thumb_path = os.path.join(THUMB_DIR, f"{safename}.{ext}")
                    if os.path.exists(thumb_path):
                        os.remove(thumb_path)

    newsletters = []

    for filename in os.listdir(NEWSLETTERS_INPUT):
        if filename.endswith('.docx') and not filename.startswith('~'):
            filepath = os.path.join(NEWSLETTERS_INPUT, filename)
            result = convert_newsletter(filepath)
            if result:
                # Check if there's a matching blogger post
                date_range = extract_date_range(filename)
                if date_range and date_range in blogger_posts:
                    result['title'] = blogger_posts[date_range]['title']
                    result['blogger_date'] = blogger_posts[date_range]['date']
                    print(f"  Matched blogger title: {result['title']}")

                newsletters.append(result)
                print(f"  Converted: {filename}")

    # Sort by modification time (newest first)
    newsletters.sort(key=lambda x: x['mod_time'], reverse=True)

    # Update index.html
    update_index(newsletters)
    print(f"Done! {len(newsletters)} newsletters active.")

def extract_end_date(title):
    """Extract the end date from a newsletter title like 'Market Summary 3-2 to 3-20 (26)'."""
    # Check for year suffix like (26) meaning 2026
    year_match = re.search(r'\((\d{2})\)\s*$', title)
    if year_match:
        year = 2000 + int(year_match.group(1))
    else:
        year = 2025  # Default year

    # Extract the end date (second date in "X-X to X-X" pattern)
    date_match = re.search(r'(\d{1,2})-(\d{1,2})\s+to\s+(\d{1,2})-(\d{1,2})', title)
    if date_match:
        end_month = int(date_match.group(3))
        end_day = int(date_match.group(4))
        try:
            end_date = datetime(year, end_month, end_day)
            return end_date.strftime("%b %d, %Y")
        except:
            pass

    return None

def update_index(newsletters):
    """Update the newsletter list in index.html."""
    with open(INDEX_FILE, 'r') as f:
        content = f.read()

    items = []
    for nl in newsletters:
        # Try to extract end date from title first
        extracted_date = extract_end_date(nl['title'])
        if extracted_date:
            date_str = extracted_date
        elif 'blogger_date' in nl:
            date_str = nl['blogger_date']
        else:
            date_str = time.strftime("%b %d, %Y", time.localtime(nl['mod_time']))
        thumb_str = f'"{nl["thumb"]}"' if nl['thumb'] else 'null'
        items.append(f'            {{ date: "{date_str}", title: "{nl["title"]}", file: "{nl["file"]}", thumb: {thumb_str} }}')

    new_array = "        newsletters = [\n" + ",\n".join(items) + "\n        ];"

    pattern = r'newsletters = \[[\s\S]*?\];'
    content = re.sub(pattern, new_array, content)

    with open(INDEX_FILE, 'w') as f:
        f.write(content)

class NewsletterHandler(FileSystemEventHandler):
    def on_any_event(self, event):
        if event.is_directory:
            return

        src = event.src_path
        if src.endswith('.docx') and not os.path.basename(src).startswith('~'):
            time.sleep(1)  # Wait for file operations to complete

            if event.event_type == 'deleted':
                print(f"\nFile deleted: {os.path.basename(src)}")
            else:
                print(f"\nChange detected: {os.path.basename(src)}")

            rebuild_all()

if __name__ == "__main__":
    print("=" * 50)
    print("Newsletter Auto-Update Watcher")
    print("=" * 50)
    print(f"Watching: {NEWSLETTERS_INPUT}")
    print("")
    print("• Add/edit a .docx file → auto-converts to website")
    print("• Delete a .docx file → removes from website")
    print("")
    print("Press Ctrl+C to stop\n")

    # Initial build
    rebuild_all()

    # Start watching
    event_handler = NewsletterHandler()
    observer = Observer()
    observer.schedule(event_handler, NEWSLETTERS_INPUT, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\nStopped watching.")

    observer.join()
