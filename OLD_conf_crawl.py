import os

import requests

import browsercookie

from urllib.parse import urlparse, parse_qs

def extract_page_id(url):

    """Extracts the pageId from a Confluence URL."""

    parsed_url = urlparse(url)
    query_params = parse_qs(parsed_url.query)
    return query_params.get("pageId", [None])[0]

def export_confluence_pages_to_pdf(urls, output_dir="confluence_pdfs"):

    """Exports a list of Confluence pages to PDF using browser cookies."""

    os.makedirs(output_dir, exist_ok=True)

    cookies = browsercookie.load()  # Load cookies from your browser

    for url in urls:

        page_id = extract_page_id(url)

        if not page_id:

            print(f"Could not extract pageId from: {url}")

            continue

        base_url = f"{urlparse(url).scheme}://{urlparse(url).netloc}"

        export_url = f"{base_url}/spaces/flyingpdf/pdfpageexport.action?pageId={page_id}"

        try:

            response = requests.get(export_url, cookies=cookies)

            if response.status_code == 200:

                filename = os.path.join(output_dir, f"page_{page_id}.pdf")

                with open(filename, "wb") as f:

                    f.write(response.content)

                print(f"✅ Exported: {filename}")

            else:

                print(f"❌ Failed to export {url} (status code: {response.status_code})")

        except Exception as e:

            print(f"⚠️ Error exporting {url}: {e}")

# Example usage

confluence_urls = [
    "https://ilabs-capco.atlassian.net/wiki/spaces/UT/pages/3737583757/7.+Architecture",
#    "https://www.bbc.co.uk",

    # Add more URLs here

]

export_confluence_pages_to_pdf(confluence_urls)