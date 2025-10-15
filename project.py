import os
import json
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
from io import BytesIO

# --------- Configure ----------
OUTPUT_DIR = "output"
IMAGE_DIR = os.path.join(OUTPUT_DIR, "images")
os.makedirs(IMAGE_DIR, exist_ok=True)
# -----------------------------

def get_image_description(image_path, use_ollama=False):
    """
    Replace this placeholder with your real LLaMA/Ollama call.
    If use_ollama=True and ollama is installed and running, this will try to use it.
    """
    if use_ollama:
        try:
            import ollama
            resp = ollama.chat(
                model="llama3.2-vision",
                messages=[{
                    "role": "user",
                    "content": "Describe this image in one sentence. Include any visible text.",
                    "images": [image_path]
                }]
            )
            return resp["message"]["content"]
        except Exception as e:
            return f"[Error getting description from Ollama: {e}]"
    # default placeholder
    return f"Auto-generated description for {os.path.basename(image_path)}"

def expand_merged_cells(ws):
    """
    Return a dict mapping (r,c) 0-based -> value,
    where merged cell top-left value is propagated to all cells in the merged range.
    """
    cell_map = {}
    # Fill with direct values first
    max_row = ws.max_row
    max_col = ws.max_column
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            val = ws.cell(row=r, column=c).value
            cell_map[(r-1, c-1)] = val

    # Now propagate merged ranges
    for merged in ws.merged_cells.ranges:
        # merged is a MergedCellRange object, e.g. 'A1:C1'
        min_row, min_col, max_row, max_col = merged.min_row, merged.min_col, merged.max_row, merged.max_col
        top_left_val = ws.cell(row=min_row, column=min_col).value
        for r in range(min_row-1, max_row):
            for c in range(min_col-1, max_col):
                cell_map[(r, c)] = top_left_val
    return cell_map, ws.max_row, ws.max_column

def image_anchor_position(image, default_row, default_col):
    """
    Return (row_index_0based, col_index_0based) of the image anchor.
    Works with different anchor shapes used by openpyxl.
    If anchor not available, returns provided defaults.
    """
    anchor = getattr(image, 'anchor', None)
    # Many openpyxl anchors have attributes ._from or .from_
    try:
        if anchor is None:
            return default_row, default_col
        # prefer anchor._from (older/newer implementations)
        from_attr = getattr(anchor, '_from', None) or getattr(anchor, 'from_', None)
        if from_attr:
            row = getattr(from_attr, 'row', None)
            col = getattr(from_attr, 'col', None)
            if row is not None and col is not None:
                return int(row), int(col)
        # fallback: anchor has .row/.col
        row = getattr(anchor, 'row', None) or getattr(anchor, 'row_coord', None)
        col = getattr(anchor, 'col', None) or getattr(anchor, 'col_coord', None)
        if row is not None and col is not None:
            return int(row), int(col)
    except Exception:
        pass
    return default_row, default_col

def is_table_block(block_rows):
    """
    Heuristic: treat a block as table if:
      - it contains >=2 rows AND
      - at least one row in the block has >=2 non-empty cells
    This handles complex tables (including merged headers).
    """
    if len(block_rows) < 2:
        return False
    for row in block_rows:
        non_empty = sum(1 for v in row if v is not None and str(v).strip() != "")
        if non_empty >= 2:
            return True
    return False

def choose_header_row(block_rows):
    """
    Choose header row index inside block_rows.
    Heuristic: pick the first row with >=2 non-empty cells; if none, pick first row.
    Returns (header_index_in_block, headers_list)
    """
    for i, row in enumerate(block_rows):
        non_empty_cells = [str(v).strip() if v is not None else "" for v in row]
        if sum(1 for v in non_empty_cells if v != "") >= 2:
            headers = [h if h not in (None, "") else f"Column_{idx+1}" for idx, h in enumerate(row)]
            return i, headers
    # fallback
    headers = [h if h not in (None, "") else f"Column_{idx+1}" for idx, h in enumerate(block_rows[0])]
    return 0, headers

def rows_to_table_dicts(headers, data_rows):
    """
    Build list of dicts mapping headers -> row values.
    If data row shorter than headers, missing values -> None.
    """
    table = []
    for row in data_rows:
        entry = {}
        for i, h in enumerate(headers):
            val = row[i] if i < len(row) else None
            entry[str(h)] = val
        table.append(entry)
    return table

def extract_sheet_ordered(ws, get_desc_fn):
    """
    Returns a list of content items preserving order.
    Each item is a dict having at least: type, order, and content fields.
    type in {"text", "table", "image"}.
    order is numeric so final sort preserves sheet order.
    """
    cell_map, max_row, max_col = expand_merged_cells(ws)
    # Build full grid as list of rows
    grid = []
    for r in range(max_row):
        row = []
        for c in range(max_col):
            row.append(cell_map.get((r, c)))
        grid.append(row)

    # collect images with positions
    images_info = []
    for img_idx, image in enumerate(ws._images):
        # default fallback position: end of sheet
        default_row = max_row
        default_col = 0
        r0, c0 = image_anchor_position(image, default_row, default_col)
        # Save image binary
        try:
            img_data = image._data()
        except Exception:
            # some openpyxl versions store .ref or .path or .image - fallback to attribute
            try:
                img_data = image.ref  # unlikely
            except Exception:
                img_data = None
        # Save file
        img_name = f"{ws.title}_image_{img_idx+1}.png"
        img_path = os.path.join(IMAGE_DIR, img_name)
        try:
            if isinstance(img_data, (bytes, bytearray)):
                img = Image.open(BytesIO(img_data))
                img.save(img_path)
            else:
                # attempt to get from image.image if present
                inner = getattr(image, 'image', None)
                if inner is not None and hasattr(inner, 'fp'):
                    inner.fp.seek(0)
                    im = Image.open(inner.fp)
                    im.save(img_path)
                else:
                    # last resort: try to save from pillow Image object if image is already PIL
                    if hasattr(image, '_data'):
                        b = image._data()
                        im = Image.open(BytesIO(b))
                        im.save(img_path)
                    else:
                        # create placeholder image so pipeline doesn't fail
                        placeholder = Image.new("RGB", (100, 40), color=(200,200,200))
                        placeholder.save(img_path)
        except Exception:
            # create placeholder
            placeholder = Image.new("RGB", (100, 40), color=(200,200,200))
            placeholder.save(img_path)

        # description via model
        description = get_desc_fn(img_path)

        images_info.append({
            "type": "image",
            "path": img_path,
            "description": description,
            "anchor_row": int(r0),
            "anchor_col": int(c0),
            # compute order: row*10000 + col to preserve order within row
            "order": int(r0) * 10000 + int(c0)
        })

    # Build ordered content by scanning rows and injecting images where anchors match row
    content = []
    r = 0
    num_rows = len(grid)
    # Build quick lookup of images by anchor row
    images_by_row = {}
    for img in images_info:
        images_by_row.setdefault(img["anchor_row"], []).append(img)

    while r < num_rows:
        # Insert images anchored at this row (in order by column)
        if r in images_by_row:
            images_here = sorted(images_by_row[r], key=lambda x: x["anchor_col"])
            for img in images_here:
                content.append({
                    "type": "image",
                    "path": img["path"],
                    "description": img["description"],
                    "order": img["order"]
                })
        # If current row empty (all cells None/blank), skip and increment
        row_vals = grid[r]
        if all((v is None or str(v).strip() == "") for v in row_vals):
            r += 1
            continue

        # Otherwise gather a block until a blank row or an image anchor row (so images appear in right place)
        block_rows = []
        block_start = r
        while r < num_rows:
            # stop block if this row has an image anchored BEFORE its data should appear
            if r in images_by_row:
                # If image column is before where text would start we might want to break; but to keep simple,
                # we break block on image row so image appears separately at same row.
                break
            row_vals = grid[r]
            if all((v is None or str(v).strip() == "") for v in row_vals):
                break
            block_rows.append(row_vals)
            r += 1

        # classify block as table or text
        if is_table_block(block_rows):
            # choose header row inside block (index)
            header_idx_in_block, headers = choose_header_row(block_rows)
            # data rows = block_rows after header row
            data_rows = block_rows[header_idx_in_block+1:]
            if not data_rows:
                # if no data rows treat header row as one-row table (still produce empty table)
                data_rows = []
            table_dicts = rows_to_table_dicts(headers, data_rows)
            content.append({
                "type": "table",
                "data": table_dicts,
                "headers": headers,
                "order": block_start * 10000
            })
        else:
            # treat each non-empty cell in the block as text content, preserve row order
            for i, brow in enumerate(block_rows):
                text = next((str(c) for c in brow if c is not None and str(c).strip() != ""), None)
                if text:
                    content.append({
                        "type": "text",
                        "content": text,
                        "order": (block_start + i) * 10000
                    })
        # continue loop (r already moved to after block)
    # It's possible some images anchor beyond last data row - append them
    extra_images = [img for img in images_info if img["anchor_row"] >= num_rows]
    for img in sorted(extra_images, key=lambda x: x["order"]):
        content.append({
            "type": "image",
            "path": img["path"],
            "description": img["description"],
            "order": img["order"]
        })

    # finally sort content by order and return
    content_sorted = sorted(content, key=lambda x: x["order"])
    return content_sorted

def extract_xlsx_ordered(filepath, get_desc_fn=get_image_description):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    workbook_data = {}
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        sheet_ordered = extract_sheet_ordered(ws, get_desc_fn)
        workbook_data[sheetname] = sheet_ordered
    return workbook_data

def save_as_json_and_md(workbook_data, base_filename):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    json_path = os.path.join(OUTPUT_DIR, f"{base_filename}.json")
    md_path = os.path.join(OUTPUT_DIR, f"{base_filename}.md")

    # Save JSON
    with open(json_path, "w", encoding="utf-8") as jf:
        json.dump(workbook_data, jf, indent=2, ensure_ascii=False)

    # Save MD
    with open(md_path, "w", encoding="utf-8") as mf:
        for sheet, items in workbook_data.items():
            mf.write(f"# Sheet: {sheet}\n\n")
            for item in items:
                if item["type"] == "text":
                    mf.write(f"- {item['content']}\n\n")
                elif item["type"] == "table":
                    headers = item.get("headers", [])
                    rows = item.get("data", [])
                    if headers:
                        mf.write("| " + " | ".join(map(str, headers)) + " |\n")
                        mf.write("| " + " | ".join(["---"] * len(headers)) + " |\n")
                        for r in rows:
                            mf.write("| " + " | ".join(str(r.get(h, "")) if r.get(h, "") is not None else "" for h in headers) + " |\n")
                        mf.write("\n")
                elif item["type"] == "image":
                    path = item.get("path")
                    desc = item.get("description", "")
                    mf.write(f"![{os.path.basename(path)}]({path})\n\n")
                    mf.write(f"**Description:** {desc}\n\n")
            mf.write("\n---\n\n")
    return json_path, md_path

# ---------------------- Usage ----------------------
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python advanced_xlsx_ordered_extractor.py <path_to_xlsx> [--ollama]")
        sys.exit(1)
    filepath = sys.argv[1]
    use_ollama_flag = ("--ollama" in sys.argv)
    def desc_fn(p): return get_image_description(p, use_ollama=use_ollama_flag)

    base = os.path.splitext(os.path.basename(filepath))[0]
    workbook_data = extract_xlsx_ordered(filepath, get_desc_fn=desc_fn)
    jpath, mpath = save_as_json_and_md(workbook_data, base)
    print("Saved:", jpath, mpath)
