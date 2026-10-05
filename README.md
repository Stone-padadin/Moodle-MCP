# Moodle MCP

A local MCP server for reading Moodle course information and downloading course files to the user's computer. It does not submit assignments or modify Moodle content.

## Requirements

- Python 3.10 or newer
- A Moodle site that enables Web Services, REST access, and the functions required by the tools
- A personal Moodle Web Services token with appropriate read permissions
- An MCP-compatible client

Moodle administrators control whether Web Services, REST, token creation, and individual functions are available. This project does not bypass school SSO or Moodle permissions. If a site only offers interactive SSO and no API token access, this connector will not work without a separate site-specific authentication integration.

## Install

```powershell
git clone https://github.com/Stone-padadin/Moodle-MCP.git
cd Moodle-MCP
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

On macOS or Linux, activate with `source .venv/bin/activate` instead.

## Configure credentials locally

Create a JSON file outside this repository:

- Windows: `%APPDATA%\moodle-mcp\config.json`
- macOS/Linux: `~/.config/moodle-mcp/config.json`

Contents:

```json
{
  "base_url": "https://moodle.example.edu",
  "token": "YOUR_PERSONAL_MOODLE_WEBSERVICE_TOKEN"
}
```

Use the base Moodle URL only: no login page path, query string, or embedded credentials. Keep this file private and do not commit or share it. Alternatively, set `MOODLE_MCP_CONFIG` to a different local config path. Set `MOODLE_MCP_DOWNLOAD_DIR` to change the local download folder; the default is `~/Downloads/Moodle`.

Create the token through your Moodle site's supported account/admin workflow. Do not paste it into chat, issues, screenshots, or source code. The token is used only by the local server to call that configured Moodle site.

## Configure your MCP client

The server uses stdio. Point your MCP client at the absolute path to `server.py` and use the Python interpreter where `requirements.txt` was installed. For example, a Codex `config.toml` entry on Windows:

```toml
[mcp_servers.moodle]
command = "C:\\path\\to\\Moodle-MCP\\.venv\\Scripts\\python.exe"
args = ["C:\\path\\to\\Moodle-MCP\\server.py"]
```

Restart the client after changing its MCP configuration.

## Tools

- `list_courses`: courses visible to the configured account
- `get_course_outline`: course sections, activities, and file metadata
- `list_course_resources`: list downloadable files, optionally filtered by extension or section
- `download_course_resources`: preview or batch-download up to 100 selected files, capped at 500 MB per batch; local existing files are never overwritten
- `read_course_resource`: extract text from PDF, PPTX, DOCX, and TXT resources
- `get_course_assignments`: assignment descriptions and open/due/cutoff timestamps
- `get_upcoming_deadlines`: Moodle calendar action events

Downloads are local filesystem writes only. File access is limited to HTTPS URLs on the configured Moodle host, individual files are capped at 25 MB, and batch downloads default to preview mode. The server exposes no Moodle write or submission tools.

## Tests

```powershell
python -m unittest discover -s tests -v
```

Tests use mocks and temporary directories; they do not require a real Moodle account or token.

## Limitations

Moodle Web Services vary by institution. A site may disable REST, token creation, or one or more API functions. APIs may not expose every item visible in the browser, and calendar events do not replace reading assessment documents. Authentication here is a user-managed Web Services token, not a universal SSO connector.
