from _csv import writer
from os import path
from os.path import dirname, basename
from subprocess import run
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askopenfilenames

import docx2txt
from PyPDF2.pdf import PdfFileReader

if __name__ == '__main__':
    root = Tk()
    root.withdraw()
    keywords_filename = askopenfilename(initialdir='.', title='Keywords File', filetypes=(('Text Files', '*.txt'),), parent=root)
    if not keywords_filename:
        exit()
    process_filenames = askopenfilenames(initialdir='.', title='File To Check', filetypes=(('PDF Files', '*.pdf'), ('Docx Files', '*.docx')), parent=root)
    if not process_filenames:
        exit()
    with open(keywords_filename, 'rt') as keywords_file:
        keywords = [keyword[:-1].upper() for keyword in keywords_file.readlines()]
    if not keywords:
        exit()
    csv_filename = path.join(dirname(keywords_filename), 'scan.csv')
    with open(csv_filename, 'wt', newline='') as csv_file:
        csv_writer = writer(csv_file)
        csv_writer.writerow([''] + keywords)
        for process_filename in process_filenames:
            if process_filename[-4:] == '.pdf':
                with open(process_filename, 'rb') as process_file:
                    pdf = PdfFileReader(process_file)
                    process_file_text = '\n'.join(pdf.getPage(i).extractText().upper() for i in range(pdf.getNumPages()))
            else:
                process_file_text = docx2txt.process(process_filename).upper()
            csv_writer.writerow([basename(process_filename)] + [False if keyword not in process_file_text else True for keyword in keywords])
    run(['start', csv_filename], shell=True)
