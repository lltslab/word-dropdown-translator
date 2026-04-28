import csv
import io
import os
import re
import tempfile
import zipfile
from pathlib import Path

from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from lxml import etree as ET

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-later")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = {"w": W_NS}
XML_PARSER = ET.XMLParser(remove_blank_text=False, recover=False)


def qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def iter_word_xml_names(zip_file: zipfile.ZipFile):
    for name in sorted(zip_file.namelist()):
        if re.fullmatch(r"word/[^/]+\.xml", name):
            yield name


def get_attr(elem, attr_name: str) -> str:
    return elem.get(qn(attr_name), "") if elem is not None else ""


def parse_xml_from_bytes(data: bytes):
    return ET.fromstring(data, parser=XML_PARSER)


def sanitize_xml_text(text: str) -> str:
    if not text:
        return text
    return "".join(
        ch for ch in text
        if ch in "\t\n\r" or ord(ch) >= 0x20
    )


def find_dropdown_controls_in_root(root):
    controls = []

    for sdt in root.findall(".//w:sdt", namespaces=XML_NS):
        sdt_pr = sdt.find("w:sdtPr", namespaces=XML_NS)
        if sdt_pr is None:
            continue

        dropdown = sdt_pr.find("w:dropDownList", namespaces=XML_NS)
        if dropdown is None:
            continue

        tag_el = sdt_pr.find("w:tag", namespaces=XML_NS)
        alias_el = sdt_pr.find("w:alias", namespaces=XML_NS)

        tag = get_attr(tag_el, "val")
        alias = get_attr(alias_el, "val")

        items = []
        for item in dropdown.findall("w:listItem", namespaces=XML_NS):
            items.append(
                {
                    "element": item,
                    "displayText": get_attr(item, "displayText"),
                    "value": get_attr(item, "value"),
                }
            )

        controls.append(
            {
                "tag": tag,
                "alias": alias,
                "items": items,
            }
        )

    return controls


def export_dropdowns_to_bytes(input_docx_path: str) -> bytes:
    rows = []
    global_control_index = 0

    with zipfile.ZipFile(input_docx_path, "r") as zin:
        for part_name in iter_word_xml_names(zin):
            data = zin.read(part_name)
            root = parse_xml_from_bytes(data)
            controls = find_dropdown_controls_in_root(root)

            for control in controls:
                for item_index, item in enumerate(control["items"], start=1):
                    rows.append(
                        {
                            "control_index": global_control_index,
                            "part": part_name,
                            "tag": control["tag"],
                            "alias": control["alias"],
                            "item_index": item_index,
                            "displayText": item["displayText"],
                            "value": item["value"],
                            "translated_displayText": "",
                        }
                    )
                global_control_index += 1

    if not rows:
        raise ValueError("No dropdown list items found in this DOCX.")

    output = io.StringIO()
    writer = csv.DictWriter(
        output,
        fieldnames=[
            "control_index",
            "part",
            "tag",
            "alias",
            "item_index",
            "displayText",
            "value",
            "translated_displayText",
        ],
    )
    writer.writeheader()
    writer.writerows(rows)

    return output.getvalue().encode("utf-8-sig")


def load_translations(csv_path: str):
    translations = {}

    with open(csv_path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        required = {"control_index", "part", "item_index", "displayText", "translated_displayText"}
        missing = required - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"CSV is missing required columns: {', '.join(sorted(missing))}")

        for row in reader:
            key = (
                row["part"],
                int(row["control_index"]),
                int(row["item_index"]),
            )
            translations[key] = {
                "original": row["displayText"],
                "translated": sanitize_xml_text((row["translated_displayText"] or "").strip()),
            }

    return translations


def import_dropdowns_to_bytes(input_docx_path: str, translated_csv_path: str):
    translations = load_translations(translated_csv_path)

    input_buffer = io.BytesIO()
    output_buffer = io.BytesIO()

    with open(input_docx_path, "rb") as f:
        input_buffer.write(f.read())
    input_buffer.seek(0)

    global_control_index = 0
    updated_count = 0

    with zipfile.ZipFile(input_buffer, "r") as zin, zipfile.ZipFile(output_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for zip_info in zin.infolist():
            data = zin.read(zip_info.filename)

            if re.fullmatch(r"word/[^/]+\.xml", zip_info.filename):
                root = parse_xml_from_bytes(data)
                changed = False

                for sdt in root.findall(".//w:sdt", namespaces=XML_NS):
                    sdt_pr = sdt.find("w:sdtPr", namespaces=XML_NS)
                    if sdt_pr is None:
                        continue

                    dropdown = sdt_pr.find("w:dropDownList", namespaces=XML_NS)
                    if dropdown is None:
                        continue

                    list_items = dropdown.findall("w:listItem", namespaces=XML_NS)
                    for item_index, list_item in enumerate(list_items, start=1):
                        key = (zip_info.filename, global_control_index, item_index)
                        if key not in translations:
                            continue

                        translated = translations[key]["translated"]
                        if translated:
                            list_item.set(qn("displayText"), translated)
                            changed = True
                            updated_count += 1

                    global_control_index += 1

                if changed:
                    xml_bytes = ET.tostring(
                        root,
                        encoding="utf-8",
                        xml_declaration=True,
                        pretty_print=False,
                        standalone=False,
                    )
                    zout.writestr(zip_info, xml_bytes)
                else:
                    zout.writestr(zip_info, data)
            else:
                zout.writestr(zip_info, data)

    output_buffer.seek(0)
    return output_buffer.read(), updated_count


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/export", methods=["POST"])
def export_route():
    file = request.files.get("docx_file")
    if not file or not file.filename.lower().endswith(".docx"):
        flash("Please upload a valid .docx file.")
        return redirect(url_for("index"))

    temp_path = None

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            file.save(tmp.name)
            temp_path = tmp.name

        csv_bytes = export_dropdowns_to_bytes(temp_path)
        output_name = Path(file.filename).stem + "_dropdown_items.csv"

        return send_file(
            io.BytesIO(csv_bytes),
            as_attachment=True,
            download_name=output_name,
            mimetype="text/csv",
        )

    except Exception as e:
        flash(str(e))
        return redirect(url_for("index"))

    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)


@app.route("/import", methods=["POST"])
def import_route():
    docx_file = request.files.get("docx_file")
    csv_file = request.files.get("csv_file")

    if not docx_file or not docx_file.filename.lower().endswith(".docx"):
        flash("Please upload a valid original .docx file.")
        return redirect(url_for("index"))

    if not csv_file or not csv_file.filename.lower().endswith(".csv"):
        flash("Please upload a valid translated .csv file.")
        return redirect(url_for("index"))

    temp_docx_path = None
    temp_csv_path = None

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            docx_file.save(tmp_docx.name)
            temp_docx_path = tmp_docx.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp_csv:
            csv_file.save(tmp_csv.name)
            temp_csv_path = tmp_csv.name

        output_bytes, updated_count = import_dropdowns_to_bytes(temp_docx_path, temp_csv_path)
        output_name = Path(docx_file.filename).stem + "_translated.docx"

        return send_file(
            io.BytesIO(output_bytes),
            as_attachment=True,
            download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        flash(str(e))
        return redirect(url_for("index"))

    finally:
        for p in [temp_docx_path, temp_csv_path]:
            if p and os.path.exists(p):
                os.remove(p)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=True)