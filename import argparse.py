import argparse
import csv
import mimetypes
import re
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import unquote, urlparse
from urllib.request import Request, urlopen

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.oxml.ns import qn


URL_PATTERN = re.compile(
    r"https?://[^\s<>\]\[\"']+",
    re.IGNORECASE
)


def category_slug(category):
    slug = re.sub(
        r"[^a-zA-Z0-9]+",
        "_",
        category.strip().lower()
    )
    return slug.strip("_") or "uncategorized"


def iter_document_blocks(document):
    """Read paragraphs and tables in document order."""

    for block in document.element.body.iterchildren():

        if isinstance(block, CT_P):
            yield Paragraph(block, document)

        elif isinstance(block, CT_Tbl):
            yield Table(block, document)


def get_paragraph_text(paragraph, document):
    """Get paragraph text and actual hyperlink URLs."""

    text = paragraph.text

    for hyperlink in paragraph._p.xpath(".//w:hyperlink"):
        relationship_id = hyperlink.get(qn("r:id"))

        if relationship_id:
            relationship = document.part.rels.get(
                relationship_id
            )

            if relationship:
                text += " " + relationship.target_ref

    return text


def read_links(file_path):
    """Extract image URLs from a DOCX file."""

    records = []
    category = "uncategorized"
    category_counts = {}

    document = Document(str(file_path))
    line_number = 0

    for block in iter_document_blocks(document):

        if isinstance(block, Paragraph):
            lines = [get_paragraph_text(block, document)]

        else:
            lines = []

            for row in block.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        lines.append(
                            get_paragraph_text(paragraph, document)
                        )

        for line in lines:
            line_number += 1

            if ":" in line and not URL_PATTERN.search(line):
                heading = line.strip().rstrip(":").strip()

                if heading:
                    category = heading

            for match in URL_PATTERN.finditer(line):

                url = match.group(0).rstrip(".,;:)")

                category_counts[category] = (
                    category_counts.get(category, 0) + 1
                )

                records.append({
                    "source_file": file_path.name,
                    "category": category,
                    "label": category_slug(category),
                    "index": category_counts[category],
                    "url": url,
                    "source_line": line_number
                })

    return records


def get_image_extension(url, content_type):

    extension = Path(
        unquote(urlparse(url).path)
    ).suffix.lower()

    valid_extensions = {
        ".jpg", ".jpeg", ".png", ".gif",
        ".webp", ".bmp", ".tif", ".tiff"
    }

    if extension in valid_extensions:
        return ".jpg" if extension == ".jpeg" else extension

    if content_type:
        extension = mimetypes.guess_extension(
            content_type.split(";")[0].strip()
        )

        if extension in valid_extensions:
            return ".jpg" if extension == ".jpeg" else extension

    return ".jpg"


def download_image(url, destination, timeout):

    request = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 ImageDatasetBuilder/1.0"
        }
    )

    try:
        with urlopen(request, timeout=timeout) as response:

            content_type = response.headers.get(
                "Content-Type", ""
            )

            if not content_type.lower().startswith("image/"):
                return "error", (
                    f"Expected an image; received "
                    f"{content_type or 'unknown content type'}"
                )

            image_data = response.read()

            if not image_data:
                return "error", "Empty response"

            destination = destination.with_suffix(
                get_image_extension(url, content_type)
            )

            destination.parent.mkdir(
                parents=True,
                exist_ok=True
            )

            destination.write_bytes(image_data)

            return "downloaded", str(destination)

    except (HTTPError, URLError, TimeoutError, OSError) as error:
        return "error", str(error)


def main():

    parser = argparse.ArgumentParser(
        description="Download categorized images from multiple DOCX files."
    )

    parser.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help="One or more DOCX input files"
    )

    parser.add_argument(
        "--output",
        type=Path,
        default=Path("image_dataset")
    )

    parser.add_argument(
        "--timeout",
        type=int,
        default=30
    )

    args = parser.parse_args()

    for input_file in args.inputs:
        if not input_file.is_file():
            parser.error(f"File not found: {input_file}")

        if input_file.suffix.lower() != ".docx":
            parser.error(f"Only DOCX files are supported: {input_file}")

    records = []

    for input_file in args.inputs:

        print(f"\nReading: {input_file.name}")

        try:
            file_records = read_links(input_file)
            records.extend(file_records)

            print(f"Found {len(file_records)} URLs")

        except Exception as error:
            parser.error(
                f"Could not read {input_file.name}: {error}"
            )

    if not records:
        parser.error("No web links found in the DOCX files.")

    args.output.mkdir(
        parents=True,
        exist_ok=True
    )

    manifest_path = args.output / "manifest.csv"

    columns = [
        "source_file",
        "category",
        "label",
        "index",
        "url",
        "source_line",
        "status",
        "local_path",
        "error"
    ]

    downloaded = 0
    failed = 0

    with manifest_path.open(
        "w",
        newline="",
        encoding="utf-8-sig"
    ) as csv_file:

        writer = csv.DictWriter(
            csv_file,
            fieldnames=columns
        )

        writer.writeheader()

        for record in records:

            filename = (
                f"{category_slug(Path(record['source_file']).stem)}_"
                f"{record['index']:04d}"
            )

            destination = (
                args.output
                / "images"
                / category_slug(record["source_file"])
                / record["label"]
                / filename
            )

            status, result = download_image(
                record["url"],
                destination,
                args.timeout
            )

            output_record = {
                **record,
                "status": status,
                "local_path": (
                    result if status == "downloaded" else ""
                ),
                "error": (
                    result if status == "error" else ""
                )
            }

            writer.writerow(output_record)

            if status == "downloaded":
                downloaded += 1
                print(
                    f"Downloaded: {record['category']} "
                    f"#{record['index']}"
                )

            else:
                failed += 1
                print(
                    f"Failed: {record['category']} "
                    f"#{record['index']}: {result}"
                )

            time.sleep(0.2)

    print("\n========== COMPLETED ==========")
    print(f"Total URLs: {len(records)}")
    print(f"Downloaded: {downloaded}")
    print(f"Failed: {failed}")
    print(f"Output folder: {args.output.resolve()}")
    print(f"CSV manifest: {manifest_path.resolve()}")


if __name__ == "__main__":
    main()