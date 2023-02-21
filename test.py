from datetime import datetime
from openpyxl import load_workbook
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
from tkinter import filedialog
# from multiprocessing.dummy import Pool

# roster_and_grades = filedialog.askopenfile()
# save_directory = filedialog.askdirectory()
# this is so we don't have to enter it every time
 
def main():
    roster_and_grades = select_wb()
    save_directory = select_save() 
    doc = select_doc('Certificate of Training - SBA Edit.docx')
    create_docs(roster_and_grades, save_directory, doc)

# selects the workbook
def select_wb():
    return (filedialog.askopenfile()).name
    # return 'BMRA Roster and Grades - 11023.0001.xlsx'

# selects save directory
def select_save():
    return filedialog.askdirectory()
    # return 'Certs'

# selects the Cert Template
def select_doc(template_file_path):
    return(template_file_path)

# creates the docs w/ correct names and converts them to PDF
def create_docs(roster_and_grades, save_directory, doc):
    doc = DocxTemplate(doc)
    wb = load_workbook(roster_and_grades)
    ws = wb.active

    certificate_number = 1
    for row in ws['2:35']:

        # Data scraped from the excel
        course_code = str(row[0].value)
        first_name = row[2].value
        last_name = row[4].value
        # grade = row[5].value
        course_name = row[6].value
        start_date = (row[7].value)
        end_date = (row[8].value)
        # agency = row[9].value
        location = row[10].value
        # class_hours = row[11].value
        clps = row[12].value

        if (isinstance(start_date, datetime)):
            start_date = start_date.strftime('%m/%d/%Y')
            end_date = end_date.strftime('%m/%d/%Y')

        # Creates a dictionary from the excel and allows for rendering of template
        template_fill = {}
        if(type(first_name) == str):
            if (end_date == start_date):
                template_fill = {
                'FNAME': first_name,
                'LNAME': last_name,
                'COURSE': course_name,
                'CLPS': clps,
                'LOCATION': location,
                'START_DATE': start_date,
            }
            else:
                end_date = '- ' + end_date
                template_fill = {
                    'FNAME': first_name,
                    'LNAME': last_name,
                    'COURSE': course_name,
                    'CLPS': clps,
                    'LOCATION': location,
                    'START_DATE': start_date,
                    'END_DATE': end_date
                }
            # creates the docs
            doc.render(template_fill)
            document_name = save_directory + "/" + str(certificate_number) + " - Certificate of Training " + course_code + ".docx"
            doc.save(document_name)
            certificate_number += 1
    # conversion()
    convert(save_directory)
    docxs = get_docx(save_directory)
    for document in docxs:
        os.remove(document)
    wb.close()

# my sad attempt at multithreaded processing
# Not used
# def conversion():
#     pool = Pool()
#     docxs = get_docx()
#     pool.map(convert, docxs) # IDK WHAT'S WRONG!!
#     for document in docxs:
#         os.remove(document)
#     pool.close()
#     pool.join()
#     print('Done!')

# creates an interable of the docx files in folder
def get_docx(save_directory):
    return (os.path.join(save_directory, file)
    for file in os.listdir(save_directory)
    if 'docx' in file)

if __name__ == "__main__":
    main()
