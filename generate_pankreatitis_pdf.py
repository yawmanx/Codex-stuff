from __future__ import annotations
from pathlib import Path
from textwrap import wrap

SOURCE = Path('sop-akute-pankreatitis.md')
OUTPUT = Path('SOP_Akute_Pankreatitis.pdf')

FONT_MAP = {
    'Helvetica': 'F1',
    'Helvetica-Bold': 'F2',
    'Courier': 'F3',
}

PAGE_WIDTH = 595  # A4 width in points
PAGE_HEIGHT = 842
LEFT_MARGIN = 72
RIGHT_MARGIN = 72
TOP_MARGIN = 72
BOTTOM_MARGIN = 72
MAX_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN

class LineEntry:
    __slots__ = ('text', 'font', 'size', 'leading')

    def __init__(self, text: str, font: str, size: float, leading: float | None = None) -> None:
        self.text = text
        self.font = font
        self.size = size
        self.leading = leading if leading is not None else size + 4


def parse_markdown() -> list[dict]:
    entries: list[dict] = []
    in_code = False
    for raw in SOURCE.read_text(encoding='utf-8').splitlines():
        if raw.strip().startswith('```'):
            in_code = not in_code
            entries.append({'type': 'blank'})
            continue
        if in_code:
            entries.append({'type': 'code', 'text': raw.rstrip('\n')})
            continue
        stripped = raw.strip()
        if not stripped:
            entries.append({'type': 'blank'})
        elif raw.startswith('# '):
            entries.append({'type': 'h1', 'text': stripped[2:].strip()})
        elif raw.startswith('## '):
            entries.append({'type': 'h2', 'text': stripped[3:].strip()})
        elif raw.lstrip().startswith('- '):
            indent = len(raw) - len(raw.lstrip())
            level = indent // 2
            entries.append({'type': 'bullet', 'text': stripped[2:].strip(), 'level': level})
        elif raw.lstrip().startswith('  - '):
            indent = len(raw) - len(raw.lstrip())
            level = indent // 2
            entries.append({'type': 'bullet', 'text': stripped[2:].strip(), 'level': level})
        elif raw.startswith('  - '):
            indent = len(raw) - len(raw.lstrip())
            level = indent // 2
            entries.append({'type': 'bullet', 'text': stripped[2:].strip(), 'level': level})
        elif raw.startswith('```'):
            continue
        elif raw.startswith('  '):
            entries.append({'type': 'text', 'text': stripped})
        else:
            entries.append({'type': 'text', 'text': stripped})
    return entries


def wrap_text(text: str, width: int) -> list[str]:
    if not text:
        return ['']
    return wrap(text, width=width)


def build_line_entries(entries: list[dict]) -> list[LineEntry]:
    result: list[LineEntry] = []
    pending_blank = False
    for entry in entries:
        etype = entry.get('type')
        if etype == 'blank':
            pending_blank = True
            continue
        if pending_blank and result:
            result.append(LineEntry('', 'Helvetica', 11, leading=14))
            pending_blank = False
        if etype == 'h1':
            text_lines = wrap_text(entry['text'], 60)
            for idx, line in enumerate(text_lines):
                content = line.upper()
                result.append(LineEntry(content, 'Helvetica-Bold', 16, leading=22 if idx == 0 else 18))
            result.append(LineEntry('', 'Helvetica', 11, leading=12))
        elif etype == 'h2':
            text_lines = wrap_text(entry['text'], 70)
            for idx, line in enumerate(text_lines):
                result.append(LineEntry(line, 'Helvetica-Bold', 13, leading=18 if idx == 0 else 16))
            result.append(LineEntry('', 'Helvetica', 11, leading=10))
        elif etype == 'code':
            # Preserve indentation; do not wrap
            result.append(LineEntry(entry['text'], 'Courier', 10, leading=13))
        elif etype == 'bullet':
            level = entry.get('level', 0)
            indent_spaces = ' ' * (4 * level)
            prefix = f"{indent_spaces}- "
            base_width = 80 - len(indent_spaces)
            wrapped = wrap_text(entry['text'], base_width)
            for idx, line in enumerate(wrapped):
                if idx == 0:
                    text = prefix + line
                else:
                    text = f"{indent_spaces}  {line}"
                result.append(LineEntry(text, 'Helvetica', 11, leading=15))
        elif etype == 'text':
            wrapped = wrap_text(entry['text'], 90)
            for line in wrapped:
                result.append(LineEntry(line, 'Helvetica', 11, leading=15))
        else:
            wrapped = wrap_text(entry.get('text', ''), 90)
            for line in wrapped:
                result.append(LineEntry(line, 'Helvetica', 11, leading=15))
    return result


def escape_pdf_text(text: str) -> str:
    return text.replace('\\', r'\\').replace('(', r'\(').replace(')', r'\)')


def paginate(lines: list[LineEntry]) -> list[str]:
    pages: list[str] = []
    current: list[str] = []
    y = PAGE_HEIGHT - TOP_MARGIN
    min_y = BOTTOM_MARGIN
    for entry in lines:
        if not entry.text:
            y -= entry.leading
            if y < min_y:
                pages.append('\n'.join(current))
                current = []
                y = PAGE_HEIGHT - TOP_MARGIN
            continue
        if y < min_y:
            pages.append('\n'.join(current))
            current = []
            y = PAGE_HEIGHT - TOP_MARGIN
        safe_text = escape_pdf_text(entry.text)
        resource = FONT_MAP[entry.font]
        command = f"BT /{resource} {entry.size} Tf {LEFT_MARGIN} {y:.2f} Td ({safe_text}) Tj ET"
        current.append(command)
        y -= entry.leading
        if y < min_y:
            pages.append('\n'.join(current))
            current = []
            y = PAGE_HEIGHT - TOP_MARGIN
    if current:
        pages.append('\n'.join(current))
    return pages


def build_pdf(pages: list[str]) -> bytes:
    parts: list[bytes] = []
    offsets: list[int] = []

    def add_object(content: bytes) -> None:
        offsets.append(sum(len(part) for part in parts))
        obj_number = len(offsets)
        header = f"{obj_number} 0 obj\n".encode('ascii')
        parts.append(header + content + b"\nendobj\n")

    # Placeholder for header
    parts.append(b"%PDF-1.4\n")
    offsets.append(len(parts[0]))  # For object 1 placeholder; will adjust later

    # We'll rebuild using a cleaner approach
    parts.clear()
    offsets.clear()
    parts.append(b"%PDF-1.4\n")

    def add(obj_str: str) -> None:
        byte_str = obj_str.encode('latin-1')
        offsets.append(sum(len(p) for p in parts))
        obj_number = len(offsets)
        parts.append(f"{obj_number} 0 obj\n".encode('ascii') + byte_str + b"\nendobj\n")

    num_pages = len(pages)
    # Catalog and pages will be added after page/content objects are known.

    # We'll collect page object numbers/content numbers to reference later.
    page_object_numbers = []
    content_object_numbers = []

    # Font objects
    add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")  # Object 1
    add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")  # Object 2
    add("<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>")  # Object 3

    font_obj_start = 1

    for page_text in pages:
        content_stream = page_text.encode('latin-1')
        stream = b"<< /Length " + str(len(content_stream)).encode('ascii') + b" >>\nstream\n" + content_stream + b"\nendstream"
        add(stream.decode('latin-1'))
        content_obj_num = len(offsets)
        content_object_numbers.append(content_obj_num)

        page_dict = (
            "<< /Type /Page /Parent {parent} 0 R /MediaBox [0 0 {w} {h}] "
            "/Resources << /Font << /F1 {f1} 0 R /F2 {f2} 0 R /F3 {f3} 0 R >> >> "
            "/Contents {contents} 0 R >>"
        ).format(
            parent=0,  # placeholder
            w=PAGE_WIDTH,
            h=PAGE_HEIGHT,
            f1=font_obj_start,
            f2=font_obj_start + 1,
            f3=font_obj_start + 2,
            contents=content_obj_num,
        )
        add(page_dict)
        page_obj_num = len(offsets)
        page_object_numbers.append(page_obj_num)

    # Pages object referencing kids
    kids = ' '.join(f"{num} 0 R" for num in page_object_numbers)
    pages_dict = f"<< /Type /Pages /Kids [{kids}] /Count {num_pages} >>"
    add(pages_dict)
    pages_obj_num = len(offsets)

    # Update page objects with correct parent reference
    updated_parts: list[bytes] = [parts[0]]
    current_obj = 1
    for obj_bytes in parts[1:]:
        obj_header, rest = obj_bytes.split(b"\n", 1)
        obj_number = int(obj_header.split()[0])
        if obj_number in page_object_numbers:
            rest_body = rest
            prefix = f"{obj_number} 0 obj\n".encode('ascii')
            body = rest_body
            placeholder = f"/Parent 0 0 R".encode('ascii')
            replacement = f"/Parent {pages_obj_num} 0 R".encode('ascii')
            body = body.replace(placeholder, replacement)
            updated_parts.append(prefix + body)
        else:
            updated_parts.append(obj_bytes)
    parts = updated_parts

    # Catalog object referencing pages
    catalog_dict = f"<< /Type /Catalog /Pages {pages_obj_num} 0 R >>"
    add(catalog_dict)
    catalog_obj_num = len(offsets)

    # Build xref
    pdf_body = b''.join(parts)
    xref_offset = len(pdf_body)

    xref_entries = [b"0000000000 65535 f \n"]
    running = 0
    for part in parts:
        xref_entries.append(f"{running:010} 00000 n \n".encode('ascii'))
        running += len(part)

    xref = b"xref\n0 " + str(len(xref_entries)).encode('ascii') + b"\n" + b''.join(xref_entries)
    trailer = (
        b"trailer\n<< /Size " + str(len(xref_entries)).encode('ascii') +
        b" /Root " + str(catalog_obj_num).encode('ascii') + b" 0 R >>\nstartxref\n" +
        str(xref_offset).encode('ascii') + b"\n%%EOF\n"
    )

    return pdf_body + xref + trailer


def main() -> None:
    entries = parse_markdown()
    lines = build_line_entries(entries)
    pages = paginate(lines)
    pdf_bytes = build_pdf(pages)
    OUTPUT.write_bytes(pdf_bytes)


if __name__ == '__main__':
    main()
