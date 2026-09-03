# SharePoint Python Downloader

Download files from an authenticated SharePoint/OneDrive session using the SharePoint REST API and Python.

This project is designed for situations where you already have access to a SharePoint file through a web browser and want to retrieve the underlying file programmatically without setting up Microsoft Graph, an Entra application, or a separate browser profile.

The downloader returns the raw file contents as `bytes`, leaving file-specific processing to the caller.

## Features

* Download files directly from SharePoint using the REST API
* Authenticate using an existing SharePoint browser session
* Uses `FedAuth` and `rtFa` session cookies
* Returns raw `bytes` without writing an intermediate file
* Works with arbitrary SharePoint files, not just Excel workbooks
* Can be combined with `pandas`, `openpyxl`, or other Python libraries
* Keeps authentication and file processing as separate concerns

## How it works

The project uses SharePoint's `GetFileByServerRelativePath` REST endpoint:

```text
/_api/Web/GetFileByServerRelativePath(decodedurl='<file_path>')/$value
```

The workflow is:

```text
Browser
   │
   │ authenticated SharePoint session
   ▼
FedAuth + rtFa
   │
   ▼
Python requests.Session()
   │
   ▼
SharePoint REST API
   │
   ▼
raw file bytes
   │
   ├── pandas
   ├── openpyxl
   ├── zipfile
   └── other file processors
```

The `$value` endpoint returns the actual file rather than the SharePoint/Office viewer page.

## Installation

Install the required packages:

```bash
pip install requests python-dotenv
```

For Excel processing with pandas:

```bash
pip install pandas openpyxl
```

Or, if using `requirements.txt`:

```text
requests
python-dotenv
pandas
openpyxl
```

## Configuration

Authentication values should **not** be hardcoded into the source code.

Create a local `.env` file:

```dotenv
SITE=https://your-tenant-my.sharepoint.com
DOMAIN=your-tenant-my.sharepoint.com

FEDAUTH=<your-FedAuth-cookie>
RTFA=<your-rtFa-cookie>

SITE_PATH=/personal/<user>
FILE_PATH=/personal/<user>/Documents/<file>.xlsx
```

Then load the configuration:

```python
import os
from dotenv import load_dotenv

load_dotenv()

SITE = os.environ["SITE"]
DOMAIN = os.environ["DOMAIN"]

FEDAUTH = os.environ["FEDAUTH"]
RTFA = os.environ["RTFA"]

SITE_PATH = os.environ["SITE_PATH"]
FILE_PATH = os.environ["FILE_PATH"]
```

### `.gitignore`

Make sure `.env` is never committed:

```gitignore
.env
*.xlsx
*.xls
*.xlsm
.ipynb_checkpoints/
__pycache__/
```

A `.env.example` file can safely be committed:

```dotenv
SITE=https://your-tenant-my.sharepoint.com
DOMAIN=your-tenant-my.sharepoint.com

FEDAUTH=<your-FedAuth-cookie>
RTFA=<your-rtFa-cookie>

SITE_PATH=/personal/<user>
FILE_PATH=/personal/<user>/Documents/<file>.xlsx
```
> How to find FedAuth-cookie?
> * Opne browser devtools and go over your webview access page
> * In the `Application` tab > click add domain contains target cookies, then copy the whole cookies value
## Usage

### Download a file

```python
import requests


def download_sharepoint_file(
    *,
    site: str,
    domain: str,
    fedauth: str,
    rtfa: str,
    site_path: str,
    file_path: str,
    timeout: int = 30,
) -> bytes:
    """
    Download a file from SharePoint using an authenticated
    SharePoint browser session.

    Returns
    -------
    bytes
        Raw contents of the requested file.
    """

    site = site.rstrip("/")
    site_path = "/" + site_path.strip("/")

    if not file_path.startswith("/"):
        file_path = "/" + file_path

    rest_url = (
        f"{site}{site_path}"
        "/_api/Web/GetFileByServerRelativePath("
        f"decodedurl='{file_path}'"
        ")/$value"
    )

    session = requests.Session()

    session.cookies.set(
        "FedAuth",
        fedauth,
        domain=domain,
        path="/",
    )

    session.cookies.set(
        "rtFa",
        rtfa,
        domain=domain,
        path="/",
    )

    response = session.get(
        rest_url,
        timeout=timeout,
        allow_redirects=False,
    )

    response.raise_for_status()

    return response.content
```

Then:

```python
file_bytes = download_sharepoint_file(
    site=SITE,
    domain=DOMAIN,
    fedauth=FEDAUTH,
    rtfa=RTFA,
    site_path=SITE_PATH,
    file_path=FILE_PATH,
)
```

The function does not write anything to disk.

## Loading an Excel workbook

For an `.xlsx` file, the returned bytes can be passed directly to pandas:

```python
import io
import pandas as pd

df = pd.read_excel(
    io.BytesIO(file_bytes),
    sheet_name="Sheet1",
)

print(df.head())
```

Only the requested worksheet is loaded into the DataFrame.

The entire XLSX file still needs to be downloaded because SharePoint's `$value` endpoint returns the workbook package as a whole.

## Working with openpyxl

If more detailed workbook manipulation is required:

```python
from io import BytesIO
from openpyxl import load_workbook

workbook = load_workbook(BytesIO(file_bytes))

print(workbook.sheetnames)

worksheet = workbook["Sheet1"]

for row in worksheet.iter_rows(values_only=True):
    print(row)
```

## Finding the required cookies

The `FedAuth` and `rtFa` values come from the authenticated browser session.

In Chromium-based browsers, they can generally be inspected through the browser's developer tools under the relevant SharePoint domain.

For this project, the important point is that the cookies must belong to the SharePoint host used by the REST request.

Do **not** paste cookie values into source code that will be committed to Git.

Do **not** print the cookie values for debugging.

## Why use the REST endpoint?

A normal SharePoint/OneDrive file link may look like:

```text
https://...sharepoint.com/:x:/r/personal/.../Documents/file.xlsx?...
```

That URL is a **viewer/share URL**. Requesting it programmatically may return an HTML Excel/Office viewer page.

The REST endpoint instead requests the underlying file:

```text
https://...sharepoint.com/personal/.../_api/Web/GetFileByServerRelativePath(decodedurl='...')/$value
```

This produces the actual file contents.

For example, an XLSX response should begin with the ZIP/XLSX signature:

```text
PK\x03\x04
```

## Security considerations

### Treat `FedAuth` and `rtFa` as credentials

These cookies represent an authenticated browser session. They should be treated as secrets.

Never commit:

```text
FEDAUTH=<real value>
RTFA=<real value>
```

to Git.

Never include them in:

* GitHub repositories
* screenshots
* notebook output
* log files
* error reports
* public issue tickets
* documentation

### `.env` is not encryption

Using `python-dotenv` does not encrypt the credentials.

It simply keeps secrets out of the source code and makes it easier to exclude them from version control.

The important protection is:

```gitignore
.env
```

### If credentials are exposed

If a real `FedAuth` or `rtFa` value is accidentally committed or publicly exposed, assume the session credential has been compromised.

Remove the secret from the repository **and invalidate/revoke the affected session** rather than relying on the credential eventually expiring.

## Limitations

This project deliberately uses the authenticated browser session rather than implementing Microsoft Graph authentication.

Consequently:

* The cookies must come from a valid authenticated session.
* The session may expire.
* The approach depends on SharePoint's authentication/session behavior.
* The REST endpoint downloads the file as a whole.
* It does not provide Microsoft Graph API permissions or application-level authentication.
* It is intended primarily for personal/local automation rather than unattended production services.

For long-running production applications, Microsoft Graph or another supported application authentication mechanism may be more appropriate.

## Example project structure

```text
sharepoint-python-downloader/
│
├── .env.example
├── .gitignore
├── README.md
├── requirements.txt
│
├── sharepoint.py
│
└── notebooks/
    └── example.ipynb
```

A typical workflow is:

```python
from io import BytesIO

import pandas as pd

from sharepoint import download_sharepoint_file

file_bytes = download_sharepoint_file(
    site=SITE,
    domain=DOMAIN,
    fedauth=FEDAUTH,
    rtfa=RTFA,
    site_path=SITE_PATH,
    file_path=FILE_PATH,
)

df = pd.read_excel(
    BytesIO(file_bytes),
    sheet_name="Sheet1",
)
```

## License

Choose an appropriate license for your project, for example MIT, Apache-2.0, or GPL-3.0.
