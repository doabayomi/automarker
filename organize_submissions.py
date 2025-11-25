#!/usr/bin/env python3
"""
organize_and_convert.py

- Uses LibreOffice headless to convert .docx -> .pdf
- Matches files to CSV rows requiring at least two name-token matches (or a fuzzy fallback)
- Extracts .zip/.rar, flattens nested folders
- Places files in OUTPUT_DIR/Index_Surname_FirstName
- If target folder exists, it won't re-copy existing files; it will only convert docx files
  that lack a corresponding pdf (no re-conversion).
"""

import os
import re
import csv
import shutil
import zipfile
import subprocess
from difflib import SequenceMatcher

# Optional: rarfile usage if available
try:
    import rarfile
    RARFILE_AVAILABLE = True
except Exception:
    rarfile = None
    RARFILE_AVAILABLE = False

# === CONFIG (uppercase as requested) ===
INPUT_DIR = "submissions"               # change as needed
OUTPUT_DIR = "organized_submissions"    # change as needed
CSV_FILE = "submitters.csv"               # change as needed
# =======================================

# --------- Utilities ---------
def normalize_text(s: str) -> str:
    """Lowercase, replace non-alphanum with space, collapse spaces."""
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def tokens_from_name_field(field: str):
    """Break a name field into tokens (handles hyphens, multiword names)."""
    return [t for t in normalize_text(field).split() if t]

def sanitize_folder_part(s: str) -> str:
    """Make a filesystem-safe single token for folder names."""
    s = (s or "").strip()
    s = re.sub(r'[^A-Za-z0-9_-]', '_', s)
    s = s or "unknown"
    return s

def safe_move(src, dst):
    """Move file, handling name collisions by adding numeric suffix."""
    base, ext = os.path.splitext(os.path.basename(dst))
    dirpath = os.path.dirname(dst)
    candidate = dst
    i = 1
    while os.path.exists(candidate):
        candidate = os.path.join(dirpath, f"{base}_{i}{ext}")
        i += 1
    shutil.move(src, candidate)
    return candidate

def safe_copy(src, dst):
    """Copy file, handling collisions by adding numeric suffix."""
    base, ext = os.path.splitext(os.path.basename(dst))
    dirpath = os.path.dirname(dst)
    candidate = dst
    i = 1
    while os.path.exists(candidate):
        candidate = os.path.join(dirpath, f"{base}_{i}{ext}")
        i += 1
    shutil.copy2(src, candidate)
    return candidate

# --------- CSV load ---------
def load_csv_rows(csv_path):
    rows = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for idx, row in enumerate(reader, start=1):
            # Keep original index if present; otherwise store a fallback _csv_index
            if not row.get("index"):
                row["_csv_index"] = str(idx)
            rows.append(row)
    return rows

# --------- Matching logic ---------
def count_token_matches(filename_tokens, row):
    """Count how many distinct name tokens (surname/first/middle) appear in filename tokens."""
    tokens = []
    for field in ("surname", "first_name", "middle_name"):
        if row.get(field):
            tokens.extend(tokens_from_name_field(row.get(field)))
    tokens = list(dict.fromkeys(tokens))  # unique while preserving order
    match_count = 0
    for t in tokens:
        if t and t in filename_tokens:
            match_count += 1
    return match_count

def fuzzy_ratio(a, b):
    return SequenceMatcher(None, a, b).ratio()

def find_best_row_for_filename(filename, rows, min_token_matches=2, fuzzy_threshold=0.60):
    """
    1) Try to find a row where at least min_token_matches tokens from that row's name appear as whole tokens in filename.
    2) If none found, fallback to fuzzy matching on normalized full name variants.
    """
    norm_fname = normalize_text(filename)
    fname_tokens = set(norm_fname.split())

    # 1) token-match phase
    best_token_match = None
    best_token_count = 0
    for row in rows:
        cnt = count_token_matches(fname_tokens, row)
        if cnt > best_token_count:
            best_token_count = cnt
            best_token_match = row

    if best_token_count >= min_token_matches:
        return best_token_match, "token"

    # 2) fuzzy fallback
    best_ratio = 0.0
    best_row = None
    for row in rows:
        # build variants: "surname first middle", "first surname"
        parts = []
        for field in ("surname", "first_name", "middle_name"):
            val = (row.get(field) or "").strip()
            if val:
                parts.append(val)
        full = normalize_text(" ".join(parts))
        rev = " ".join(full.split()[::-1]) if full else ""
        for variant in [full, rev]:
            if not variant:
                continue
            r = fuzzy_ratio(norm_fname, variant)
            if r > best_ratio:
                best_ratio = r
                best_row = row

    if best_ratio >= fuzzy_threshold:
        return best_row, f"fuzzy({best_ratio:.2f})"

    return None, None

# --------- Archive extraction & flattening ---------
def extract_and_flatten_archive(archive_path, extract_to):
    os.makedirs(extract_to, exist_ok=True)
    if archive_path.lower().endswith(".zip"):
        try:
            with zipfile.ZipFile(archive_path, "r") as zf:
                zf.extractall(extract_to)
        except Exception as e:
            print(f"Error extracting zip {archive_path}: {e}")
    elif archive_path.lower().endswith(".rar"):
        if not RARFILE_AVAILABLE:
            print(f"⚠️ rarfile not available; skipping extraction for {archive_path}")
            return
        try:
            with rarfile.RarFile(archive_path, "r") as rf:
                rf.extractall(extract_to)
        except rarfile.RarCannotExec:
            print(f"⚠️ No 'unrar' tool found. Install 'unrar' or skip .rar files. Skipping {archive_path}")
            return
        except Exception as e:
            print(f"Error extracting rar {archive_path}: {e}")
            return

    # flatten nested directories into extract_to
    for root, dirs, files in os.walk(extract_to):
        for f in files:
            src = os.path.join(root, f)
            dst = os.path.join(extract_to, f)
            if os.path.abspath(src) != os.path.abspath(dst):
                # handle collisions
                if os.path.exists(dst):
                    # create a unique name
                    base, ext = os.path.splitext(f)
                    i = 1
                    while True:
                        candidate = os.path.join(extract_to, f"{base}_{i}{ext}")
                        if not os.path.exists(candidate):
                            dst = candidate
                            break
                        i += 1
                shutil.move(src, dst)
    # remove empty subdirectories
    for root, dirs, _ in os.walk(extract_to, topdown=False):
        for d in dirs:
            try:
                os.rmdir(os.path.join(root, d))
            except OSError:
                pass

# --------- DOCX -> PDF conversion (LibreOffice) ---------
def libreoffice_convert_docx_to_pdf(docx_path, out_dir):
    """
    Convert docx to pdf using LibreOffice headless.
    Places the converted PDF into out_dir.
    Returns path to pdf if success else None.
    """
    if not shutil.which("libreoffice"):
        print("⚠️ 'libreoffice' not found in PATH. Install it (e.g., sudo apt install libreoffice).")
        return None

    # Use --outdir to put the pdf into the target folder reliably
    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except subprocess.CalledProcessError as e:
        print(f"Failed converting {docx_path} with LibreOffice: {e}")
        return None

    pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    return pdf_path if os.path.exists(pdf_path) else None

# --------- Main processing ---------
def organize_and_convert():
    rows = load_csv_rows(CSV_FILE)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    stats = {
        "files_seen": 0, "matched": 0, "unmatched": 0,
        "converted": 0, "convert_failed": 0, "skipped_existing_file": 0
    }

    for fname in sorted(os.listdir(INPUT_DIR)):
        if fname.startswith("."):
            continue
        stats["files_seen"] += 1
        input_path = os.path.join(INPUT_DIR, fname)
        if not os.path.isfile(input_path):
            continue

        matched_row, how = find_best_row_for_filename(fname, rows, min_token_matches=2, fuzzy_threshold=0.60)
        if not matched_row:
            print(f"No good match found for: {fname}")
            stats["unmatched"] += 1
            continue

        # get index (prefer explicit CSV "index" column, else fallback to _csv_index)
        index = matched_row.get("index") or matched_row.get("_csv_index") or ""
        surname = matched_row.get("surname") or ""
        first_name = matched_row.get("first_name") or ""
        folder_name = f"{index}_{sanitize_folder_part(surname)}_{sanitize_folder_part(first_name)}"
        target_folder = os.path.join(OUTPUT_DIR, folder_name)
        os.makedirs(target_folder, exist_ok=True)

        # If file isn't already in the target folder, copy it there
        dest_path = os.path.join(target_folder, os.path.basename(input_path))
        if os.path.exists(dest_path):
            print(f"File already present in target folder; skipping copy: {fname}")
            stats["skipped_existing_file"] += 1
        else:
            # If archive -> extract into target_folder; else copy file
            if fname.lower().endswith((".zip", ".rar")):
                # extract into temporary subfolder then flatten into target_folder
                extract_and_flatten_archive(input_path, target_folder)
            else:
                safe_copy(input_path, dest_path)

        # Now convert any .docx in target_folder that lack a corresponding .pdf
        for root, _, files in os.walk(target_folder):
            for f in files:
                if f.lower().endswith(".docx"):
                    docx_full = os.path.join(root, f)
                    pdf_expected = os.path.splitext(docx_full)[0] + ".pdf"
                    if os.path.exists(pdf_expected):
                        # already converted — skip
                        continue
                    pdf_path = libreoffice_convert_docx_to_pdf(docx_full, root)
                    if pdf_path:
                        stats["converted"] += 1
                        # Optionally remove the original docx (comment/uncomment as desired)
                        # os.remove(docx_full)
                    else:
                        stats["convert_failed"] += 1

        stats["matched"] += 1
        print(f"Processed '{fname}' → folder '{folder_name}' (match method: {how})")

    # Summary
    print("\n=== SUMMARY ===")
    for k, v in stats.items():
        print(f"{k}: {v}")


if __name__ == "__main__":
    organize_and_convert()
