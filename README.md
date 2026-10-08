# PdfPasswordRemover

Removes the protection (printing / copying / editing restrictions) set on PDF files.

The script processes every PDF in a source directory and creates an unprotected copy of each one in a destination directory. Originals are kept as a backup.

> Only PDFs that can be opened without a password are supported (i.e. protected by an *owner* password only). PDFs that ask for a password to be opened are reported as failed.

## Installation

### Windows executable

Download `PdfPasswordRemover.exe` from the [latest release](https://github.com/doolieSoft/PdfPasswordRemover/releases/latest). No Python installation is needed.

### From source

Requires Python 3.8+.

```bash
pip install -r requirements.txt
```

## Usage

```bash
python PdfPasswordRemover.py <source_directory> <destination_directory> [-a SECONDS] [-p]
```

| Argument | Description |
|---|---|
| `source_directory` | Directory containing the protected PDFs |
| `destination_directory` | Directory where the results are written |
| `-a`, `--age_of_file_to_treat` | Only treat files older than this number of seconds (useful when files are still being written into the source directory) |
| `-p`, `--pause` | Wait for the user to press enter before exiting (useful when launched by double-click) |
| `-v`, `--version` | Show the version and exit |

### Example

```bash
python PdfPasswordRemover.py C:\Inbox C:\Output -a 60
```

## Behaviour

Each run creates two timestamped directories in `destination_directory`:

- `YYYYMMDDHHMMSS_new`: the unprotected PDFs
- `YYYYMMDDHHMMSS_backup`: the original files, moved out of the source directory

Details:

- Files ending with `.pdf`, `.pdf.convert` or `.pdf.import` (case-insensitive) are treated. The `.convert` / `.import` suffix is removed from the output name, and the original case of the file name is kept.
- Other files are ignored and stay in the source directory.
- If a PDF cannot be treated, it is left in the source directory and the error is displayed.

## Releasing a new version

1. Update `__version__` in `PdfPasswordRemover.py` and commit.
2. Tag the commit and push the tag:

   ```bash
   git tag -a v2.0.1 -m "Version 2.0.1"
   git push origin master v2.0.1
   ```

GitHub Actions then builds `PdfPasswordRemover.exe` and publishes it in a GitHub release. The build fails if the tag does not match `__version__`.
