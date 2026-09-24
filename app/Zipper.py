
import zipfile

# ===== SETTINGS =====
FOLDER = r"C:\path\to\your\zip\folder"   # change this
SEARCH_SUBFOLDERS = True      # True = also find zips inside subfolders
SEPARATE_SUBFOLDERS = False   # True = each zip goes into its own subfolder (named after the zip)
DELETE_AFTER_EXTRACT = False  # True = delete the zip after a successful extract
# ====================


def safe_extract(zf, dest):
    """Extract while blocking paths that try to escape the destination folder."""
    dest_abs = os.path.abspath(dest)
    for member in zf.infolist():
        target = os.path.abspath(os.path.join(dest, member.filename))
        if not target.startswith(dest_abs + os.sep) and target != dest_abs:
            print(f"   Skipped unsafe path: {member.filename}")
            continue
        zf.extract(member, dest)


def find_zips(folder):
    """Find zip files by content (not just extension)."""
    found = []
    if SEARCH_SUBFOLDERS:
        for root, _, files in os.walk(folder):
            for f in files:
                p = os.path.join(root, f)
                if zipfile.is_zipfile(p):
                    found.append(p)
    else:
        for f in os.listdir(folder):
            p = os.path.join(folder, f)
            if os.path.isfile(p) and zipfile.is_zipfile(p):
                found.append(p)
    return sorted(found)


def main():
    folder = os.path.abspath(FOLDER)
    if not os.path.isdir(folder):
        print(f"Folder not found: {folder}")
        return

    zips = find_zips(folder)
    total = len(zips)
    print(f"Found {total} zip files in {folder}\n")

    if total == 0:
        print("Nothing found. Here is what the script sees in this folder:")
        for f in os.listdir(folder)[:30]:
            print(f"   {f!r}")
        return

    ok, failed = 0, []

    for i, zip_path in enumerate(zips, 1):
        zip_dir = os.path.dirname(zip_path)   # extract next to the zip itself
        name = os.path.basename(zip_path)
        if SEPARATE_SUBFOLDERS:
            dest = os.path.join(zip_dir, os.path.splitext(name)[0])
            os.makedirs(dest, exist_ok=True)
        else:
            dest = zip_dir

        print(f"[{i}/{total}] Extracting {name} ...")
        try:
            with zipfile.ZipFile(zip_path) as zf:
                safe_extract(zf, dest)
            ok += 1
            if DELETE_AFTER_EXTRACT:
                os.remove(zip_path)
        except Exception as e:
            print(f"   ERROR: {e}")
            failed.append(zip_path)

    print(f"\nDone. Extracted: {ok}, Failed: {len(failed)}")
    for f in failed:
        print(f"   - {f}")


if __name__ == "__main__":
    main()
