import base64
import json
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / ".codex_work" / "extracted"
OUT.mkdir(parents=True, exist_ok=True)


def clean_text(value):
    if isinstance(value, list):
        return "".join(str(x) for x in value)
    return str(value or "")


def extract_notebook(path):
    data = json.loads(path.read_text(encoding="utf-8"))
    stem = path.stem
    image_dir = OUT / f"{stem}_images"
    image_dir.mkdir(exist_ok=True)
    lines = [f"NOTEBOOK: {path.relative_to(ROOT)}", f"CELL_COUNT: {len(data.get('cells', []))}", ""]
    summary = {"path": str(path.relative_to(ROOT)), "cells": [], "errors": [], "images": []}
    image_index = 0
    for idx, cell in enumerate(data.get("cells", []), start=1):
        cell_type = cell.get("cell_type", "unknown")
        source = clean_text(cell.get("source", []))
        execution_count = cell.get("execution_count")
        lines.extend([f"===== CELL {idx} | {cell_type} | execution_count={execution_count} =====", source.rstrip(), ""])
        record = {"index": idx, "type": cell_type, "execution_count": execution_count, "source_chars": len(source), "outputs": []}
        for output_idx, output in enumerate(cell.get("outputs", []), start=1):
            output_type = output.get("output_type", "unknown")
            record["outputs"].append(output_type)
            lines.append(f"--- OUTPUT {output_idx} | {output_type} ---")
            if output_type == "stream":
                text = clean_text(output.get("text", []))
                lines.append(text.rstrip())
            elif output_type == "error":
                error_text = "\n".join(output.get("traceback", []))
                lines.append(f"{output.get('ename')}: {output.get('evalue')}\n{error_text}")
                summary["errors"].append({"cell": idx, "name": output.get("ename"), "value": output.get("evalue")})
            else:
                payload = output.get("data", {})
                for mime, value in payload.items():
                    if mime == "image/png":
                        image_index += 1
                        image_path = image_dir / f"cell_{idx:03d}_output_{output_idx:02d}_{image_index:03d}.png"
                        image_path.write_bytes(base64.b64decode(clean_text(value)))
                        rel = image_path.relative_to(ROOT)
                        lines.append(f"[IMAGE EXTRACTED: {rel}]")
                        summary["images"].append({"cell": idx, "path": str(rel)})
                    elif mime.startswith("text/") or mime in {"application/json"}:
                        text = clean_text(value)
                        lines.append(f"[{mime}]\n{text.rstrip()}")
                    else:
                        lines.append(f"[{mime} omitted]")
            lines.append("")
        summary["cells"].append(record)
    (OUT / f"{stem}.txt").write_text("\n".join(lines), encoding="utf-8")
    (OUT / f"{stem}_summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return summary


def extract_docx(path):
    doc = Document(path)
    lines = [f"DOCUMENT: {path.relative_to(ROOT)}", f"PARAGRAPHS: {len(doc.paragraphs)}", f"TABLES: {len(doc.tables)}", ""]
    for idx, paragraph in enumerate(doc.paragraphs, start=1):
        style = paragraph.style.name if paragraph.style else ""
        text = paragraph.text
        if text.strip():
            lines.append(f"P{idx:04d} [{style}] {text}")
    for table_idx, table in enumerate(doc.tables, start=1):
        lines.extend(["", f"===== TABLE {table_idx} ({len(table.rows)}x{len(table.columns)}) ====="])
        for row in table.rows:
            lines.append(" | ".join(cell.text.replace("\n", " / ") for cell in row.cells))
    out_path = OUT / f"{path.stem}.txt"
    out_path.write_text("\n".join(lines), encoding="utf-8")
    return {"paragraphs": len(doc.paragraphs), "tables": len(doc.tables), "output": str(out_path.relative_to(ROOT))}


results = {
    "notebooks": [
        extract_notebook(ROOT / "FER" / "msc-major-project-fer.ipynb"),
        extract_notebook(ROOT / "CWD" / "msc-major-project-cwd.ipynb"),
    ],
    "document": extract_docx(ROOT / "Report" / "MSc Major Project Report Draft.docx"),
}
(OUT / "extraction_summary.json").write_text(json.dumps(results, indent=2), encoding="utf-8")
print(json.dumps({
    "notebooks": [{"path": x["path"], "cells": len(x["cells"]), "errors": x["errors"], "images": len(x["images"])} for x in results["notebooks"]],
    "document": results["document"],
}, indent=2))
