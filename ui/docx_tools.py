import re
from copy import deepcopy


PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}<>]+?)\s*\}\}|<<\s*([^{}<>]+?)\s*>>")


def normalize_placeholder(name):
    return re.sub(r"\s+", "_", name.strip()).lower()


def display_placeholder(name):
    return normalize_placeholder(name).replace("_", " ").title()


def extract_placeholders(doc):
    placeholders = []
    seen = set()

    def collect(text):
        for match in PLACEHOLDER_RE.findall(text or ""):
            raw = match[0] or match[1]
            key = normalize_placeholder(raw)
            if key and key not in seen:
                seen.add(key)
                placeholders.append(key)

    for para in doc.paragraphs:
        collect(para.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    collect(para.text)

    return placeholders


def replace_placeholders_in_text(text, values):
    def repl(match):
        key = normalize_placeholder(match.group(1) or match.group(2))
        value = values.get(key, "")
        return value if value else match.group(0)

    return PLACEHOLDER_RE.sub(repl, text or "")


def replace_placeholders_in_paragraph(para, values):
    original = para.text
    replaced = replace_placeholders_in_text(original, values)
    if original == replaced:
        return

    if not para.runs:
        para.add_run(replaced)
        return

    first_run = para.runs[0]
    first_run.text = replaced
    for run in para.runs[1:]:
        run.text = ""


def replace_placeholders(doc, values):
    normalized = {normalize_placeholder(k): v for k, v in values.items()}

    for para in doc.paragraphs:
        replace_placeholders_in_paragraph(para, normalized)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_placeholders_in_paragraph(para, normalized)

    return doc


def filled_copy(doc, values):
    copied = deepcopy(doc)
    return replace_placeholders(copied, values)
