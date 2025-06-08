# GMail .eml file importer

Wanting to go back from Proton Mail to Gmail, the only option Proton Mail offers when exporting is a folder with .eml files.
Using IMAP sync in Gmail ignores the original message date. Thankfully we can use the Gmail API instead.

## Features

- ✅ **Preserves original message dates** (unlike IMAP import)
- ✅ **Duplicate detection** based on Message-ID headers
- ✅ **Progress bar** with real-time status updates
- ✅ **Automatic label creation** and application
- ✅ **Recursive directory processing**
- ✅ **Robust error handling** with detailed statistics

# Requirements
- Python 3.6+

# Setup
Steps in order to use the importer:

# 1. The importer itself
1. Clone the repo with  `git clone git@github.com/zundrium/gmail-eml-importer.git`.
2. Create a virtual environment with `python -m venv .venv`.
3. Activate the virtual environment with `source .venv/bin/activate` on Linux or `.\.venv\Scripts\Activate.ps1` on Windows.
4. Install the packages with `pip install -r requirements.txt`.

# 2. Getting access to your own Gmail API
1. Go to https://console.cloud.google.com/.
2. Create a new project or select existing one.
3. Setup oAuth consent screen, make it external and add yourself as a test user.
4. Search for "Gmail" and Enable Gmail API.
5. Create OAuth2 credentials (left column) -> oAuth client ID (Desktop app).
6. Download the credentials.json file and place it in this folder.

# Usage
```
# Default (looks for credentials.json in current directory)
python gmail_eml_importer.py /path/to/eml/files -l "Imported"

# Specify credentials file location
python gmail_eml_importer.py /path/to/eml/files -c /path/to/credentials.json -l "Import"

# With all options
python gmail_eml_importer.py /path/to/eml/files -c credentials.json -l "Import" -r --no-duplicates
```