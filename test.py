from openpyxl import load_workbook
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
from tkinter import filedialog
from multiprocessing.dummy import Pool

roster_and_grades = filedialog.askopenfile()
save_directory = filedialog.askdirectory()

doc = DocxTemplate('Certificate of Training - Edit.docx')
wb = load_workbook(roster_and_grades.name)
ws = wb.active

def main():
    create_docs()
    # print(type(get_docx()))

def create_docs():
    certificate_number = 1
    for row in ws['2:35']: # should find a better way to do this 

        # Data scraped from the excel
        course_code = str(row[0].value)
        first_name = row[2].value
        last_name = row[4].value
        grade = row[5].value
        course_name = row[6].value
        start_date = str(row[7].value)
        end_date = str(row[8].value)
        agency = row[9].value
        location = row[10].value
        class_hours = row[11].value
        clps = row[12].value

        # Creates a JSON from the excel and allows for rendering of template
        if(type(first_name) == str):
            template_fill = {
                'FNAME': first_name,
                'LNAME': last_name,
                'COURSE': course_name,
                'CLPS': clps,
                'LOCATION': location,
                'START_DATE': start_date,
                'END_DATE': end_date
            }

            doc.render(template_fill)
            # add file paths below
            document_name = save_directory + "/" + str(certificate_number) + " Certificate of Training " + course_code + ".docx"
            doc.save(document_name)
            # convert(str(certificate_number) + " Certificate of Training " + course_code + ".docx")
            # os.remove(str(certificate_number) + " Certificate of Training " + course_code + ".docx")
            certificate_number += 1

    pool = Pool()
    docxs = get_docx()
    pool.map(convert, docxs)
    pool.close()
    pool.join()

def get_docx():
    return (os.path.join(save_directory, file)
      for file in os.listdir(save_directory)
      if 'docx' in file)

if __name__ == "__main__":
    main()