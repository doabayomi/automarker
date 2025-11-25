import os
import csv
import time
import json
import dotenv
# --- Standardized Import for google-genai SDK ---
import google.generativeai as genai
from google.generativeai import types # Use this if you need specific types, though often not needed for basic calls

dotenv.load_dotenv()

# ========== CONFIG ==========
INPUT_DIR = "organized_submissions"    # Folder containing submission subfolders
OUTPUT_FILE = "grades.csv"             # Grades file
PROMPT_FILE = "marking_prompt.txt"     # Shared prompt file
MODEL_NAME = "gemini-2.0-flash"        # Using 1.5-flash as 2.0-flash might be an older name or limited access
MAX_TOKEN_LIMIT = 300000               # Adjust based on your quota
BASE_WAIT = 30                         # Initial retry delay (seconds)
# ============================

# --- Setup Gemini (CORRECTED & STANDARDIZED) ---
# Use genai.configure() as it is correct for this import style.
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

if not GEMINI_KEY:
    raise ValueError("GEMINI_API_KEY environment variable not set.")

# Correct configuration method for 'import google.generativeai as genai'
genai.configure(api_key=GEMINI_KEY) 

# Correct model initialization method
model = genai.GenerativeModel(MODEL_NAME)


def list_folders(root):
    """List all subfolders."""
    return [d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]


def read_existing_grades(csv_path):
    """Read already processed entries."""
    seen = set()
    if os.path.exists(csv_path):
        with open(csv_path, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                key = (row.get("index"), row.get("surname"), row.get("first_name"))
                seen.add(key)
    return seen


def write_grade_row(csv_path, row, fieldnames):
    """Append one result row to the CSV file, updating header if new fields appear."""
    existing_fieldnames = []
    file_exists = os.path.exists(csv_path)

    if file_exists:
        # Read existing header to preserve column order
        with open(csv_path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            existing_fieldnames = reader.fieldnames or []

    # Merge new fieldnames
    all_fields = existing_fieldnames or []
    for f in fieldnames:
        if f not in all_fields:
            all_fields.append(f)

    # Rewrite file if header changes
    if not file_exists or set(all_fields) != set(existing_fieldnames):
        # Read existing rows (if any)
        rows = []
        if file_exists:
            with open(csv_path, newline="", encoding="utf-8") as f:
                rows = list(csv.DictReader(f))

        # Rewrite with updated header
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=all_fields)
            writer.writeheader()
            for r in rows:
                writer.writerow(r)
            writer.writerow(row)
    else:
        # Append normally
        with open(csv_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=all_fields)
            writer.writerow(row)


def upload_files(file_paths):
    """Upload files to Gemini."""
    uploaded = []
    for fpath in file_paths:
        try:
            # Correct call for this import style: genai.upload_file
            file_ref = genai.upload_file(fpath)
            uploaded.append(file_ref)
            print(f"📤 Uploaded: {os.path.basename(fpath)}")
        except Exception as e:
            print(f"❌ Upload failed for {fpath}: {e}")
    return uploaded


def delete_uploaded_files(file_refs):
    """Delete uploaded Gemini files."""
    for file_ref in file_refs:
        try:
            # Correct call for this import style: genai.delete_file
            genai.delete_file(file_ref.name)
            print(f"🗑️ Deleted: {file_ref.name}")
        except Exception as e:
            print(f"⚠️ Could not delete {file_ref.name}: {e}")


def process_folder(folder_name, seen_set, prompt_base):
    """Grade a single submission folder."""
    parts = folder_name.split("_")
    if len(parts) < 3:
        print(f"⚠️ Skipping invalid folder name: {folder_name}")
        return False

    index, surname, first_name = parts[:3]
    key = (index, surname, first_name)

    if key in seen_set:
        print(f"⏩ Skipping {folder_name} (already processed)")
        return False

    folder_path = os.path.join(INPUT_DIR, folder_name)
    has_accdb = any(f.lower().endswith(".accdb") for f in os.listdir(folder_path))

    # Collect PDFs and images
    file_paths = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith((".pdf", ".png", ".jpg", ".jpeg", ".bmp"))
    ]

    if not file_paths:
        print(f"⚠️ No PDFs or images in {folder_name}")
        return False

    uploaded_files = upload_files(file_paths)
    if not uploaded_files:
        return False

    try:
        # Build final prompt
        prompt = (
            prompt_base
            + "\n\nInclude this info in your reasoning for this submission: "
              f"has_accdb = {str(has_accdb).lower()}.\n"
            + "Respond strictly as JSON in this format:\n"
              "{\n  problem_ident: <int>,\n  erd: <int>,\n  schema: <int>,\n  conclusion: <int>\n}"
        )

        # Token estimation
        try:
            # Correct call for this import style: model.count_tokens
            token_info = model.count_tokens([prompt] + uploaded_files)
            total_tokens = token_info.total_tokens
            print(f"📏 Token count for {folder_name}: {total_tokens}")
        except Exception as e:
            print(f"❌ Token count failed for {folder_name}: {e}")
            return False

        if total_tokens > MAX_TOKEN_LIMIT:
            print(f"🚫 Skipping {folder_name}, token count too high ({total_tokens})")
            return False

        retries = 3
        wait = BASE_WAIT
        for attempt in range(1, retries + 1):
            try:
                # Correct call for this import style: model.generate_content
                response = model.generate_content(
                    [prompt] + uploaded_files,
                    generation_config={"response_mime_type": "application/json"},
                )
                text = response.text.strip()
                print(f"✅ Response for {folder_name}: {text}")

                try:
                    result = json.loads(text)
                except json.JSONDecodeError:
                    print(f"⚠️ JSON parse failed on attempt {attempt} for {folder_name}")
                    if attempt < retries:
                        time.sleep(wait)
                        wait *= 2
                        continue
                    return False

                # Save result
                row = {"index": index, "surname": surname, "first_name": first_name}
                for k, v in result.items():
                    row[k] = v
                fieldnames = ["index", "surname", "first_name"] + sorted(result.keys())
                write_grade_row(OUTPUT_FILE, row, fieldnames)
                print(f"💾 Saved {folder_name}")
                return True

            except Exception as e:
                # Handle rate limit/server error
                if "429" in str(e) or "503" in str(e):
                    print(f"⏱️ Rate limit/server error, retry {attempt}/{retries}")
                    time.sleep(wait)
                    wait *= 2
                    continue
                print(f"❌ Fatal error for {folder_name}: {e}")
                return False

        print(f"❌ All retries failed for {folder_name}")
        return False
    finally:
        delete_uploaded_files(uploaded_files)


def main():
    if not os.path.exists(PROMPT_FILE):
        print(f"Error: Missing {PROMPT_FILE}")
        return

    with open(PROMPT_FILE, "r", encoding="utf-8") as f:
        prompt_base = f.read()

    seen = read_existing_grades(OUTPUT_FILE)

    if not os.path.isdir(INPUT_DIR):
        print(f"Error: Input folder '{INPUT_DIR}' not found")
        return

    for folder in list_folders(INPUT_DIR):
        process_folder(folder, seen, prompt_base)


if __name__ == "__main__":
    main()