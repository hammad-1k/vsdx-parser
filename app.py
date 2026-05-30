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
        label = " ".join(p for p in text_parts if p)
        if label:
            shapes[sid] = label

    edges = []
    seen = set()
    for connect in root.findall(f".//{tag('Connect')}"):
        from_sheet = connect.get("FromSheet")
        to_sheet = connect.get("ToSheet")
        if not from_sheet or not to_sheet:
            continue
        key = f"{from_sheet}->{to_sheet}"
        if key in seen:
            continue
        seen.add(key)
        edges.append({
            "from_id": from_sheet,
            "to_id": to_sheet,
            "from_label": shapes.get(from_sheet, ""),
            "to_label": shapes.get(to_sheet, ""),
        })

    steps = []
    for i, sid in enumerate(shapes.keys(), start=1):
        outgoing = [e["to_label"] for e in edges if e["from_id"] == sid and e["to_label"]]
        steps.append({
            "step": i,
            "id": sid,
            "label": shapes[sid],
            "outgoing_to": outgoing,
        })

    return {
        "process_name": process_name,
        "step_count": len(steps),
        "steps": steps,
        "edges": edges,
    }

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/parse", methods=["POST"])
def parse():
    secret = os.environ.get("PARSER_SECRET", "")
    if secret:
        incoming = request.headers.get("X-Parser-Secret", "")
        if incoming != secret:
            return jsonify({"error": "Unauthorized"}), 401
    file_bytes = None
    if "file" in request.files:
        file_bytes = request.files["file"].read()
    elif request.data:
        file_bytes = request.data
    else:
        return jsonify({"error": "No file provided"}), 400
    try:
        result = parse_vsdx(file_bytes)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run()
