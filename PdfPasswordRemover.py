import argparse
import os
import shutil
import stat
import time

import PyPDF2


def main():
    parser = argparse.ArgumentParser(
        description="Tool made to remove password set on pdf files - (Stefano Crapanzano - s.crapanzano@gmail.com)")
    parser.add_argument('source_directory', help='Directory containing the PDFs with password protection')
    parser.add_argument('destination_directory', help='Directory where the new PDFs will be placed')
    parser.add_argument('-a', '--age_of_file_to_treat', help='Age in seconds of file to be treated')
    args = parser.parse_args()
    source_directory = vars(args)['source_directory']
    destination_directory = vars(args)['destination_directory']
    age_of_file_to_treat = vars(args)['age_of_file_to_treat']

    timestamp = time.localtime()
    date_exec = str(
        timestamp.tm_year) + str(timestamp.tm_mon).zfill(2) + str(timestamp.tm_mday).zfill(2) + str(
        timestamp.tm_hour).zfill(2) + str(
        timestamp.tm_min).zfill(2) + str(timestamp.tm_sec).zfill(2)

    destination_backup_folder_name = os.path.join(destination_directory, date_exec + '_backup')
    destination_new_folder_name = os.path.join(destination_directory, date_exec + '_new')

    nb_pdf_corrected = 0
    nb_files_moved = 0

    for file_name in os.listdir(source_directory):

        source_file_path = os.path.join(source_directory, file_name)

        if not os.path.isfile(source_file_path):
            continue

        if age_of_file_to_treat and file_age_in_seconds(source_file_path) < float(age_of_file_to_treat):
            continue

        try:
            if not os.path.isdir(destination_backup_folder_name):
                os.mkdir(destination_backup_folder_name)
        except:
            print("Exception occurred while creating " + destination_backup_folder_name)
            break

        try:
            if not os.path.isdir(destination_new_folder_name):
                os.mkdir(destination_new_folder_name)
        except:
            print("Exception occurred while creating " + destination_new_folder_name)
            break

        old_file_name = file_name
        file_name = file_name.lower().replace('.pdf.convert', '.pdf')
        file_name = file_name.lower().replace('.pdf.import', '.pdf')
        destination_file_path = os.path.join(destination_new_folder_name, file_name)

        if file_name.lower().endswith(('.pdf', '.pdf.convert', 'pdf.import')):
            f = None
            pdf_output_file = None

            try:
                f = open(source_file_path, 'rb')

                file_reader = PyPDF2.PdfReader(f)

                print(source_file_path + " > Nb pages: " + str(len(file_reader.pages)) + "; Is Encrypted: " + str(
                    file_reader.is_encrypted))

                pdf_writer = PyPDF2.PdfWriter()

                for pageNum in range(len(file_reader.pages)):
                    page_obj = file_reader.pages[pageNum]
                    pdf_writer.add_page(page_obj)

                pdf_output_file = open(destination_file_path, 'wb')
                pdf_writer.write(pdf_output_file)
                pdf_output_file.close()
                nb_pdf_corrected = nb_pdf_corrected + 1
                f.close()
            except:
                if f is not None:
                    f.close()

                if pdf_output_file is not None:
                    pdf_output_file.close()

                print("--- Exception occurred with creation of file " + destination_file_path + " ---")
        try:
            shutil.move(source_file_path, os.path.join(destination_backup_folder_name, old_file_name))
            nb_files_moved = nb_files_moved + 1
        except:
            print("Exception occurred while moving from [ " + source_file_path + " ] to [ " +
                  destination_backup_folder_name + old_file_name + " ]")

    print(str(nb_files_moved) + " file(s) have been moved in " + destination_backup_folder_name)
    print(str(nb_pdf_corrected) + " pdf file(s) have been treated.")
    input("Press enter to continue...")

def file_age_in_seconds(pathname):
    return time.time() - os.stat(pathname)[stat.ST_MTIME]


if __name__ == "__main__":
    main()