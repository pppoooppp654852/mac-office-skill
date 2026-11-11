"""Constants shared by the runner modules."""

import os
import tempfile

APP_NAMES = {
    "word": "Microsoft Word",
    "excel": "Microsoft Excel",
    "powerpoint": "Microsoft PowerPoint",
}

APP_BY_EXTENSION = {".docx": "word", ".xlsx": "excel", ".pptx": "powerpoint"}
APP_FOR_KIND = {extension[1:]: app for extension, app in APP_BY_EXTENSION.items()}
KIND_FOR_APP = {app: kind for kind, app in APP_FOR_KIND.items()}

GROUP_CONTAINER_DIR = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/mac-office-staging"
)

# Exported PDFs keyed by the source's sha256, so re-rendering an unchanged file never opens
# the Office app again. The per-user temp directory is purged by macOS on its own schedule.
PDF_CACHE_DIR = os.path.join(tempfile.gettempdir(), "mac-office-pdf-cache")

BUNDLE_IDS = {
    "word": "com.microsoft.Word",
    "excel": "com.microsoft.Excel",
    "powerpoint": "com.microsoft.Powerpoint",
}

DOCUMENT_COLLECTIONS = {"word": "documents", "excel": "workbooks", "powerpoint": "presentations"}

# The runner keeps per-app lock and lifecycle files here; they are reused across runs.
RUNNER_STATE_DIRNAME = "runner-state"

# An app the runner launched (hidden) is quit on a later call once it sat unused this long.
IDLE_QUIT_SECONDS = 600

# Each Office app declares `display alerts` with a different type: Word uses WdAlertLevel,
# Excel uses a boolean, and PowerPoint exposes no such property at all. The runner has
# already started the app (hidden) before this preamble runs, so it never says `launch`.
ALERT_HANDLERS = {
    "word": (
        'tell application "Microsoft Word"\n\tset display alerts to alerts none\nend tell\n',
        'tell application "Microsoft Word"\n\tset display alerts to alerts all\nend tell\n',
    ),
    "excel": (
        'tell application "Microsoft Excel"\n\tset display alerts to false\nend tell\n',
        'tell application "Microsoft Excel"\n\tset display alerts to true\nend tell\n',
    ),
    "powerpoint": (
        'tell application "Microsoft PowerPoint"\n'
        '\trepeat 20 times\n'
        '\t\ttry\n'
        '\t\t\tget Version\n'
        '\t\t\texit repeat\n'
        '\t\ton error\n'
        '\t\t\tdelay 0.1\n'
        '\t\tend try\n'
        '\tend repeat\n'
        'end tell\n',
        "",
    ),
}

# Marker the runner-owned scripts return when the app swallowed the open (it can refuse
# legacy or blocked formats without raising an AppleScript error).
OPEN_FAILED = "mac-office-open-failed"

TCC_HINT = (
    "Not authorized to send Apple events. Open System Settings > Privacy & Security > "
    "Automation and allow your terminal/IDE to control Microsoft Word, Excel, PowerPoint "
    "and System Events (used to keep the apps hidden). To re-trigger the prompt: "
    "tccutil reset AppleEvents"
)

MODAL_ALERT_HINT = (
    "Target app raised a modal dialog (such as 'Grant File Access', an unreadable-content "
    "repair prompt or a save prompt). The runner keeps the app hidden, so bring it forward "
    "from the Dock, dismiss the dialog and retry."
)

UNREADABLE_DIALOG_HINT = (
    "Timed out and dialog state could not be read (grant Accessibility permission to "
    "detect dialogs). Check the app for a blocking dialog, then retry."
)

NO_DIALOG_HINT = "No dialog found. The operation may still be running; raise --timeout."

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

CONTAINER_HINTS = {
    "ole2": (
        "OLE2 container (Word 97-2003 / legacy binary format). Only the native Office app "
        "can read or write it: unzip, LibreOffice and textutil cannot. Saving from the native "
        "app keeps the OLE2 container regardless of the file extension."
    ),
    "ole2_mip": (
        "OLE2 container holding a Microsoft Purview / RMS EncryptedPackage stream. Only the "
        "native Office app (signed in with the required rights) can read or write it: unzip, "
        "LibreOffice and textutil cannot. Saving from the native app keeps the OLE2 container "
        "regardless of the file extension."
    ),
}

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
S_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

EMU_PER_POINT = 12700

# Native-app inspectors answer with one osascript stdout blob. Field, record and section
# separators are ASCII control characters, which cannot occur in document text (unlike tab
# and newline), so no lossy escaping or re-splitting is needed.
FIELD_SEPARATOR = "\x1f"
RECORD_SEPARATOR = "\x1d"
SECTION_SEPARATOR = "\x1e"
