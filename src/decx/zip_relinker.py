"""ZIP-based OLE/chart link rewriter — modifies PPTX XML directly.

Rewrites external link paths in .rels files inside the PPTX zip,
avoiding the ~1s per-link COM overhead of LinkFormat.SourceFullName.
Run BEFORE COM opens the file.
"""

import logging
import os
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile

log = logging.getLogger(__name__)


def relink_pptx_zip(pptx_path: str, new_excel_path: str) -> int:
    """Rewrite OLE/chart link paths in the PPTX zip.

    Replaces the Excel file path in all external .rels entries with
    new_excel_path, preserving the !sheet!range suffix.

    Args:
        pptx_path: Path to the PPTX file (modified in-place).
        new_excel_path: Absolute path to the new Excel file.

    Returns count of links rewritten.
    """
    new_excel_path = os.path.abspath(new_excel_path)
    # Convert to file:/// URI format (forward slashes)
    new_file_uri = "file:///" + new_excel_path.replace("\\", "/")

    rewritten = 0

    # Work on a temp copy, then replace original
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pptx")
    os.close(tmp_fd)

    try:
        with (
            zipfile.ZipFile(pptx_path, "r") as zin,
            zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout,
        ):
            for item in zin.infolist():
                data = zin.read(item.filename)

                # Only process .rels files in slides/ and charts/
                if item.filename.endswith(".rels") and (
                    "slides/_rels/" in item.filename or "charts/_rels/" in item.filename
                ):
                    modified, count = _rewrite_rels(data, new_file_uri)
                    if count > 0:
                        data = modified
                        rewritten += count

                zout.writestr(item, data)

        # Replace original with modified
        shutil.move(tmp_path, pptx_path)
    except Exception:
        # Clean up temp file on error
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise

    if rewritten:
        log.info("ZIP relink: rewrote %d link(s) in %s", rewritten, pptx_path)

    return rewritten


def _rewrite_rels(data: bytes, new_file_uri: str) -> tuple[bytes, int]:
    """Rewrite external relationship targets in a .rels XML file.

    Returns (modified_bytes, count_of_rewritten_links).
    """
    root = ET.fromstring(data)
    ns = root.tag.split("}")[0] + "}" if "}" in root.tag else ""
    count = 0

    for rel in root:
        target_mode = rel.get("TargetMode", "")
        target = rel.get("Target", "")

        if target_mode != "External" or not target.startswith("file:///"):
            continue

        # Split target into file path and suffix (!sheet!range)
        # Target format: "file:///C:\path\data.xlsx!Tables!R1C1:R5C5"
        # or just: "file:///C:\path\data.xlsx"
        bang_pos = target.find("!", len("file:///"))
        if bang_pos >= 0:
            suffix = target[bang_pos:]  # "!Tables!R1C1:R5C5"
            new_target = new_file_uri + suffix
        else:
            new_target = new_file_uri

        if target != new_target:
            rel.set("Target", new_target)
            count += 1

    if count == 0:
        return data, 0

    # Serialize back to XML bytes
    # Preserve the XML declaration and namespace
    ET.register_namespace("", ns.strip("{}")) if ns else None
    modified = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
    return modified, count


def detect_linked_excel(pptx_path: str) -> str | None:
    """Detect the Excel file path from existing OLE links in a PPTX.

    Parses the PPTX zip for the first external OLE relationship target
    and returns the file path (stripped of file:/// prefix and !sheet!range suffix).

    Returns None if no linked Excel is found.
    """
    try:
        with zipfile.ZipFile(pptx_path, "r") as z:
            rels_files = [
                f for f in z.namelist() if "slides/_rels/" in f and f.endswith(".rels")
            ]

            for rf in sorted(rels_files):
                root = ET.fromstring(z.read(rf))
                for rel in root:
                    target = rel.get("Target", "")
                    target_mode = rel.get("TargetMode", "")
                    if target_mode != "External" or not target.startswith("file:///"):
                        continue

                    # Strip file:/// prefix
                    path = target[len("file:///") :]
                    # Strip !sheet!range suffix
                    bang_pos = path.find("!")
                    if bang_pos >= 0:
                        path = path[:bang_pos]
                    # Normalize slashes
                    path = path.replace("/", os.sep)

                    if path and os.path.exists(path):
                        return path

    except Exception:
        pass

    return None
