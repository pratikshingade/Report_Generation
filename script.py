import glob
import os
import shutil
import time

from pdf_generation.generate_pdf_statement import GeneratePdfStatement
from pdf_generation.generate_pdf_summary import GeneratePdfSummary
from pdf_generation.generate_canal_summary import GenerateCanalExcel
from summary_statement_merge_pdf import merge_pdf_reports


def statement_func(excel_path):
    pdf_instance_statement = GeneratePdfStatement(excel_path)
    pdf_instance_statement.create_pdf()


def summary_func(excel_path):
    pdf_instance_summary = GeneratePdfSummary(excel_path)
    pdf_instance_summary.create_pdf()


def canal_abstract_excel_func(folder_path):
    canal_abstract = GenerateCanalExcel(folder_path=folder_path)
    canal_abstract.create_pdf()


def copy_pdf(directory, name_of_canal):
    out_path = os.path.split(directory)[0]
    out_folder = os.path.join(out_path, f"{name_of_canal} All Pdfs")
    final_out_path = os.path.join(out_path, out_folder)
    if not os.path.exists(path=out_folder):
        os.makedirs(out_folder)
    pdf_path = glob.glob(os.path.join(directory, "*.pdf"))[0]
    shutil.copyfile(pdf_path, os.path.join(final_out_path, os.path.split(pdf_path)[1]))


def canal_name(path):
    last_folder_name = os.path.split(path)[1]
    f_list = last_folder_name.split(" ")

    i = 0
    name_list = []
    while True:
        name_list.append(f_list[i])
        i += 1
        if f_list[i][0].isdigit():
            break
    name_list.append(f_list[i])
    c_name = " ".join(name_list)
    return c_name


def main(directory, name_of_canal):

    statement_count = 0
    summary_count = 0

    for root, sub_dirs, files in os.walk(directory):
        for file in files:
            if file.startswith("Sta") and file.endswith((".xls", ".xlsx")):
                statement_count += 1
                print(file)
                file_path = os.path.join(root, file)
                statement_func(file_path)

            elif file.startswith("Sum") and file.endswith((".xls", ".xlsx")):
                summary_count += 1
                print(file)
                file_path = os.path.join(root, file)
                summary_func(file_path)

    print(f"{statement_count} Statement PDFs Are Created")
    print(f"{summary_count} Summary PDFs Are Created\n")

    merge_pdf_reports(directory, name_of_canal)
    print("All PDF's merge Files are Created\n")

    canal_abstract_excel_func(folder_path=directory)
    copy_pdf(directory=directory, name_of_canal=name_of_canal)
    print("Canal Abstract Excel and PDF File Created and Copied...!")


if __name__ == '__main__':
    start_time = time.time()
    """FOR SINGLE FILES"""
    # statement_func(r"C:\Users\ss\Downloads\Sta_बोरगाव_K_KVT-1-39.xlsx")
    # summary_func(r"C:\Users\ss\Downloads\Sum_बोरगाव_K_KVT-1-39.xlsx")

    """FOR PARTICULAR CANAL DIRECTORY WALK"""
    # walk_directory = input("Enter Walk Directory: ")
    # canal_name = canal_name(walk_directory)  # input("Enter Canal Name: ")
    # main(directory=walk_directory, name_of_canal=canal_name)

    """FOR ALL CANALS DIRECTORY WALK"""
    walk_directory = input("Enter Walk Directory For All Canals: ")

    for root, sub_dirs, files in os.walk(walk_directory):
        for sub_dir in sub_dirs:
            if sub_dir.__contains__("All Excel"):
                w_directory = os.path.join(root, sub_dir)
                cl_name = canal_name(w_directory)
                main(directory=w_directory, name_of_canal=cl_name)
    print(f"{time.time() - start_time:.2f} Secs To Complete All Canals")
