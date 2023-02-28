import os

from PyPDF2 import PdfMerger


def merge_pdf_reports(walk_directory, canal_name):
    for root, sub_dirs, files in os.walk(walk_directory):
        pdf_list = []

        for file in files:
            if file.endswith(("Statement.pdf", "Summary.pdf")):
                pdf_list.append(os.path.join(root, file))
                if len(pdf_list) > 1:
                    pdf_list.sort(reverse=True)

                    # print(pdf_list)
                    merger = PdfMerger()

                    for pdf in pdf_list:
                        merger.append(pdf)
                    out_path = os.path.split(walk_directory)[0]
                    out_folder = os.path.join(out_path, f"{canal_name} All Pdfs")
                    # print(out_folder)

                    if not os.path.exists(path=out_folder):
                        os.makedirs(out_folder)

                    merger.write(os.path.join(out_folder, " ".join(file.split(" ")[:-1]) + " Final.pdf"))
