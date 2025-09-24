# Standards Clause Extractor

Utilities for breaking a standards PDF into numbered clauses and exporting the results as JSON and Excel. A minimal web UI lets reviewers upload a PDF, browse the extracted clauses, and download the data.

## Features

- Parses headings like `1`, `1.1`, `1.1.1`, … and preserves the clause hierarchy.
- Filters boilerplate (headers, footers, licensing banners) and rebuilds readable paragraphs.
- Exports to both JSON (nested structure) and Excel (flattened table with hierarchy metadata).
- Optional web interface for uploads, inline browsing with "More" modals, and one-click downloads.

## Prerequisites

- **Python**: 3.10+ recommended (tested with 3.10.12).
- **Python dependencies**: Install the requirements once (installs `pdfminer.six`). No external command-line utilities are needed anymore.

## Project Layout

```
- Standards/                # Sample input PDF
- -output/                  # Generated outputs
- src/
  ├─ extract_clauses.py     # Core extraction logic + CLI
  └─ server.py              # Minimal HTTP UI (uses stdlib http.server)
- README.md
- requirements.txt
```

## Quick Start

### 1. Install Python dependencies

```bash
python3 -m pip install -r requirements.txt
```

### 2. Run the extractor from the CLI

```bash
python3 src/extract_clauses.py Standards/your-standard.pdf --output-dir output-folder
```

Outputs:

- `output-folder/clauses.json`: nested clause tree (`subclauses` array for children).
- `output-folder/clauses.xlsx`: Excel workbook with columns `Clause`, `Title`, `Parent`, `Level`, `Text`.

Run `python3 src/extract_clauses.py -h` for full CLI options.

### 3. Launch the web UI (optional)

Because the UI uses the Python standard library HTTP server you can run it directly:

```bash
python3 -m src.server
```

Then visit [http://127.0.0.1:8000](http://127.0.0.1:8000) and upload a PDF. The page displays the clause table and offers download links for the generated JSON and Excel (served via data URIs).

> **Note:** Binding to ports below 1024 or running on hardened environments may require additional privileges. Use a different port if needed:
>
> ```bash
> python3 -m src.server  # defaults to 127.0.0.1:8000
> ```
>
> To change the port, edit `run()` in `src/server.py` or wrap it in your own launcher.

## Deploying Elsewhere

1. Clone or copy the repository to the target machine.
2. Create a virtual environment if desired:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
3. Use the CLI or start the web UI as shown above. For production, consider placing the script behind a proper WSGI server or reverse proxy; `src/server.py` is intentionally lightweight and lacks hardening.

## Customisation Notes

- The heuristics that skip boilerplate live in `src/extract_clauses.py` (`SKIP_PATTERNS` and `looks_like_fragment`). Adjust them if your standards use different watermarks or formatting.
- Row truncation and the modal behaviour are implemented in plain JavaScript within `src/server.py`; tweak the limit or styling in `truncate_text()` / the embedded CSS.
- The extractor now parses layout information via `pdfminer.six`. Different PDFs might still require refinement—validate the outputs when onboarding new document families.

## Troubleshooting

- **`ModuleNotFoundError: pdfminer`** – run `python3 -m pip install -r requirements.txt` to install dependencies.
- **Empty outputs** – verify the PDF contains numeric headings following the expected pattern; table-heavy documents may need new parsing rules.
- **Permission errors when starting the server** – avoid privileged ports or run behind a socket handed to the process by your init system.

## License

Use within your organisation’s policy. Update this section with formal licensing if required.
