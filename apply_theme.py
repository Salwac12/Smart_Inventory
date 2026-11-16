from pathlib import Path
from datetime import datetime
import shutil

# ==========================================
# SMART INVENTORY THEME INSTALLER
# ==========================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATES_DIR = BASE_DIR / "templates"

THEME_FILE = BASE_DIR / "static" / "css" / "theme.css"

BACKUP_DIR = BASE_DIR / (
    "theme_backup_" +
    datetime.now().strftime("%Y%m%d_%H%M%S")
)

THEME_LINK = (
    '<link rel="stylesheet" '
    'href="{{ url_for(\'static\', '
    'filename=\'css/theme.css\') }}">'
)

HTML_FILES = [
    "dashboard.html",
    "inventory.html",
    "add_component.html",
    "stock.html",
    "transactions.html",
    "upload_bom.html",
    "match_bom.html",
    "calculate_materials.html",
]


def apply_theme():

    print("\nSMART INVENTORY THEME INSTALLER")
    print("--------------------------------")

    if not THEME_FILE.exists():
        print("ERROR: theme.css was not found.")
        print("Create static/css/theme.css first.")
        return

    if not TEMPLATES_DIR.exists():
        print("ERROR: templates folder not found.")
        return

    files_to_update = []

    for filename in HTML_FILES:

        file_path = TEMPLATES_DIR / filename

        if not file_path.exists():
            print(f"SKIPPED: {filename} not found")
            continue

        content = file_path.read_text(encoding="utf-8-sig")

        # Avoid adding the stylesheet twice.
        if "css/theme.css" in content:
            print(f"ALREADY CONNECTED: {filename}")
            continue

        # Find the closing head tag.
        head_position = content.lower().find("</head>")

        if head_position == -1:
            print(f"SKIPPED: No </head> in {filename}")
            continue

        # Add stylesheet AFTER existing page CSS.
        updated = (
            content[:head_position]
            + "\n    "
            + THEME_LINK
            + "\n"
            + content[head_position:]
        )

        files_to_update.append(
            (file_path, content, updated)
        )

    if not files_to_update:
        print("\nNo files require updating.")
        return

    # Backup original templates before modification.
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    for file_path, original, updated in files_to_update:

        backup_path = BACKUP_DIR / file_path.name

        shutil.copy2(file_path, backup_path)

        file_path.write_text(
            updated,
            encoding="utf-8"
        )

        print(f"UPDATED: {file_path.name}")

    print("\nTheme installation completed.")

    print(
        "Original templates backed up to:",
        BACKUP_DIR
    )

    print("\nYour database was not modified.")


if __name__ == "__main__":
    apply_theme()