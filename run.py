#
# Written by Ramtin Mesgari
# For use by Blacfox
#

import os
from tkinter import *
from tkinter import filedialog
from docx import Document
import powerpoint as pp

##############
# # FIELDS # #
##############

# TKINTER #
root = Tk()
root.minsize(500, 150)
root.title("DOCX to PPTX")

# POWERPOINT #
# file = "template.pptx"
# prs = Presentation(file)
#
# slides = prs.slides  # get slides
# slide_index = []  # slide index
# slide_idx = []  # slide id
#
# for s in slides:
#     slide_index.append(slides.index(s))
#     slide_idx.append(slides[slides.index(s)].slide_id)

# GLOBAL VARIABLES #
file_name = None

title = None
introduction = None
before_title = None
middle_title = None
after_title = None
quote = None
before_array = []
middle_array = []
after_array = []
company = None
person = None
website = None
employees = None
no_employees = None
prod_serv = None
prod_serv_list = []


#################
# # FUNCTIONS # #
#################

# def get_slide(index):
#     return slides.get(slide_idx[slide_index[index]])


# Remove any empty lines from a string
def full(text):
    return os.linesep.join([s for s in text.splitlines() if s])


def get_file_name(text):
    temp = str(text).split("/")
    temp2 = str(temp[len(temp) - 1]).split(".")

    return temp2[0]


def split_array(array):
    string = ""

    for arr in array:
        string += (arr + "\n")

    return string


def get_indices(para):
    indices = []

    for i in range(len(para)):

        if "Before:" in para[i].text:
            indices.append(i)
        elif "Why SAP Partner and Solution" in para[i].text:
            indices.append(i)
        elif "After:" in para[i].text:
            indices.append(i)
        elif '"' in para[i].text:
            indices.append(i)

    return indices


def get_bullets(para, a, b):
    bullets = []

    for i in range(a + 1, b):
        if para[i].text not in ["\n", "\r\n", ""]:
            bullets.append(full(para[i].text))

    return bullets


def get_others(para, quote_index):
    others = []

    for i in range(quote_index + 1, len(para)):
        if para[i].text not in ["\n", "\r\n", ""]:
            others.append(full(para[i].text))

    return others


def open_file(event):
    root.filename = filedialog.askopenfilename(initialdir="/Desktop/", title="Select a File",
                                               filetypes=(("Microsoft Office Word ", "*.doc"),
                                                          ("Microsoft Office Word ", "*.docx")))

    global file_name
    file_name = get_file_name(root.filename)

    lbl.config(text="You have opened " + file_name, padx=8)
    lbl.pack()

    event.pack(padx=8, pady=(10, 20))

    doc = Document(root.filename)
    para = doc.paragraphs

    global title
    title = full(para[0].text)
    global introduction
    introduction = full(para[1].text)

    before_title_index = get_indices(para)[0]
    middle_title_index = get_indices(para)[1]
    after_title_index = get_indices(para)[2]
    quote_index = get_indices(para)[3]

    global before_title
    before_title = full(para[before_title_index].text)
    global middle_title
    middle_title = full(para[middle_title_index].text)
    global after_title
    after_title = full(para[after_title_index].text)
    global quote
    quote = full(para[quote_index].text)

    global before_array
    before_array = get_bullets(para, before_title_index, middle_title_index)
    global middle_array
    middle_array = get_bullets(para, middle_title_index, after_title_index)
    global after_array
    after_array = get_bullets(para, after_title_index, quote_index)

    other_array = get_others(para, quote_index)

    global company
    company = other_array[0]
    global person
    person = other_array[1]
    global website
    website = other_array[2]
    global employees
    employees = other_array[3]
    global no_employees
    no_employees = other_array[4]
    global prod_serv
    prod_serv = other_array[5]
    global prod_serv_list
    prod_serv_list = []

    for i in range(6, len(other_array)):
            prod_serv_list.append(other_array[i])


def convert_file(event):
    pp.run()

    pp.heading.text = str(title)
    pp.introduction.text = str(introduction)

    pp.before_heading.text = str(before_title)
    pp.middle_heading.text = str(middle_title)
    pp.after_heading.text = str(after_title)

    pp.quote.text = str(quote)
    pp.percentage1.text = "XX%"
    pp.percentage_text1.text = "Percentage information"
    pp.percentage2.text = "YY%"
    pp.percentage_text2.text = "Percentage information"
    pp.customer_name_heading.text = "Customer Name"
    pp.customer_name_text.text = "Customer location\n(City, Country/State)"
    pp.industry_heading.text = "Industry"
    pp.industry_text.text = "Designated\nSAP industry"
    pp.products_and_services_heading.text = "Products and Services"
    pp.employees_heading.text = "Employees"
    pp.employees_text.text = str(no_employees)
    pp.revenue_heading.text = "Revenue"
    pp.revenue_text.text = "Insert text here\n(add US$ or € where applicable)"
    pp.featured_solutions_heading.text = "Featured Solutions"
    pp.featured_solutions_text.text = "SAP solutions (max. two solutions)"
    pp.video_heading.text = "To watch the video"
    pp.video_text.text = "Video link"

    pp.products_and_services_text.text = str(split_array(prod_serv_list))
    pp.before_text.text = str(split_array(before_array))
    pp.middle_text.text = str(split_array(middle_array))
    pp.after_text.text = str(split_array(after_array))

    pp.save(company + ".pptx")

    lbl.config(text="Your file has been converted to " + str(file_name) + ".pptx", padx=8)
    event.pack_forget()


btnOpen = Button(root, text="Open File", command=lambda: open_file(btnConvert), width=12)
btnOpen.pack(padx=8, pady=(20, 10))

lbl = Label(root)

btnConvert = Button(root, text="Convert File", command=lambda: convert_file(btnConvert), width=12)

root.mainloop()
