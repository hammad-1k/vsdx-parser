import os
import io
import zipfile
import xml.etree.ElementTree as ET
from flask import Flask, request, jsonify

app = Flask(__name__)
NS = "http://schemas.microsoft.com/office/visio/2012/main"

def tag(name):
    return f"{{{NS}}}{name}"

def parse_vsdx(file_bytes):
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        process_name = "Unknown Process"
        try:
            core_xml = zf.read("docProps/core.xml")
            core = ET.fromstring(core_xml)
            dc_ns = "http://purl.org/dc/elements/1.1/"
            title_el = core.find(f"{{{dc_ns}}}title")
            if title_el is not None and title_el.text:
                process_name = title_el.text.strip()
        except Exception:
            pass
        try:
            page_xml = zf.read("visio/pages/page1.xml")
        except KeyError:
            raise ValueError("No page1.xml found.")

    root = ET.fromstring(page_xml)
    shapes = {}
    for shape in root.findall(f".//{tag('Shape')}"):
        sid = shape.get("ID")
        if not sid:
            continue
        text_parts = []
        for t in shape.findall(f".//{tag('Text')}"):
            text_parts.append("".join(t.itertext()).strip())
        l
