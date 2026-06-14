import json
import os
import re
import shutil
import time


class TemplateLibrary:
    ROOT = os.path.join(os.getcwd(), "template_library")
    INDEX_PATH = os.path.join(ROOT, "templates.json")

    @classmethod
    def ensure(cls):
        os.makedirs(cls.ROOT, exist_ok=True)
        if not os.path.exists(cls.INDEX_PATH):
            cls._write_index([])

    @classmethod
    def _read_index(cls):
        cls.ensure()
        try:
            with open(cls.INDEX_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []

    @classmethod
    def _write_index(cls, items):
        os.makedirs(cls.ROOT, exist_ok=True)
        with open(cls.INDEX_PATH, "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)

    @staticmethod
    def slugify(name):
        slug = re.sub(r"[^a-zA-Z0-9]+", "_", name.strip()).strip("_").lower()
        return slug or f"template_{int(time.time())}"

    @classmethod
    def list_templates(cls):
        return cls._read_index()

    @classmethod
    def add_template(cls, name, source_path, placeholders=None):
        cls.ensure()
        slug = cls.slugify(name)
        filename = f"{slug}.docx"
        target_path = os.path.join(cls.ROOT, filename)
        counter = 2
        while os.path.exists(target_path):
            filename = f"{slug}_{counter}.docx"
            target_path = os.path.join(cls.ROOT, filename)
            counter += 1

        shutil.copy2(source_path, target_path)
        items = [item for item in cls._read_index() if item.get("path") != target_path]
        items.insert(0, {
            "name": name.strip() or slug.replace("_", " ").title(),
            "path": target_path,
            "placeholders": placeholders or [],
            "updated_at": time.strftime("%Y-%m-%d %H:%M"),
        })
        cls._write_index(items)
        return target_path

    @classmethod
    def save_document_as_template(cls, name, doc, placeholders=None):
        cls.ensure()
        slug = cls.slugify(name)
        target_path = os.path.join(cls.ROOT, f"{slug}.docx")
        counter = 2
        while os.path.exists(target_path):
            target_path = os.path.join(cls.ROOT, f"{slug}_{counter}.docx")
            counter += 1
        doc.save(target_path)

        items = cls._read_index()
        items.insert(0, {
            "name": name.strip() or os.path.splitext(os.path.basename(target_path))[0],
            "path": target_path,
            "placeholders": placeholders or [],
            "updated_at": time.strftime("%Y-%m-%d %H:%M"),
        })
        cls._write_index(items)
        return target_path
