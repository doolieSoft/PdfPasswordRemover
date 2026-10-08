import argparse
import shutil
import time
from pathlib import Path

from pypdf import PdfReader, PdfWriter

__version__ = '2.0.0'

# Extensions handled, with the suffix to strip to get back a plain ".pdf" name
PDF_SUFFIXES = ('.pdf', '.pdf.convert', '.pdf.import')
EXTRA_SUFFIXES = ('.convert', '.import')


def parse_args():
    parser = argparse.ArgumentParser(
        description="Tool made to remove password set on pdf files - (Stefano Crapanzano - s.crapanzano@gmail.com)")
    parser.add_argument('source_directory', type=Path, help='Directory containing the PDFs with password protection')
    parser.add_argument('destination_directory', type=Path, help='Directory where the new PDFs will be placed')
    parser.add_argument('-a', '--age_of_file_to_treat', type=float,
                        help='Minimum age in seconds of the files to treat (younger files are ignored)')
    parser.add_argument('-p', '--pause', action='store_true',
                        help='Wait for the user to press enter before exiting')
    parser.add_argument('-v', '--version', action='version', version=f'%(prog)s {__version__}')
    return parser.parse_args()


def is_pdf(path):
    return path.name.lower().endswith(PDF_SUFFIXES)


def output_name(file_name):
    """Strip the extra '.convert' / '.import' suffix, keeping the original case."""
    for suffix in EXTRA_SUFFIXES:
        if file_name.lower().endswith('.pdf' + suffix):
            return file_name[:-len(suffix)]
    return file_name


def file_age_in_seconds(path):
    return time.time() - path.stat().st_mtime


def remove_password(source_path, destination_path):
    reader = PdfReader(source_path)
    print(f"{source_path} > Nb pages: {len(reader.pages)}; Is Encrypted: {reader.is_encrypted}")

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    with open(destination_path, 'wb') as output_file:
        writer.write(output_file)


def main():
    args = parse_args()

    date_exec = time.strftime('%Y%m%d%H%M%S')
    backup_directory = args.destination_directory / f'{date_exec}_backup'
    new_directory = args.destination_directory / f'{date_exec}_new'

    nb_pdf_corrected = 0
    nb_pdf_failed = 0

    for source_path in sorted(args.source_directory.iterdir()):
        if not source_path.is_file() or not is_pdf(source_path):
            continue

        if args.age_of_file_to_treat is not None and file_age_in_seconds(source_path) < args.age_of_file_to_treat:
            continue

        try:
            backup_directory.mkdir(parents=True, exist_ok=True)
            new_directory.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            print(f"Exception occurred while creating the destination directories: {e}")
            break

        destination_path = new_directory / output_name(source_path.name)

        try:
            remove_password(source_path, destination_path)
        except Exception as e:
            nb_pdf_failed += 1
            destination_path.unlink(missing_ok=True)
            print(f"--- Exception occurred with creation of file {destination_path}: {e} ---")
            print(f"--- {source_path} is left in place ---")
            continue

        try:
            shutil.move(source_path, backup_directory / source_path.name)
            nb_pdf_corrected += 1
        except OSError as e:
            print(f"Exception occurred while moving [ {source_path} ] to [ {backup_directory} ]: {e}")

    print(f"{nb_pdf_corrected} pdf file(s) have been treated and moved in {backup_directory}")
    if nb_pdf_failed:
        print(f"{nb_pdf_failed} pdf file(s) could not be treated and were left in {args.source_directory}")

    if args.pause:
        input("Press enter to continue...")


if __name__ == "__main__":
    main()
