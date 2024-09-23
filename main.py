#This uses Spire.Doc to replace words in a Word doc
from datetime import date, datetime
import fitz
from spire.doc import *
from spire.pdf import *
import tkinter
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog as fd
import csv

ORIGINAL_NEW_UI19="documents\\New_UI-19.docx"
ORIGINAL_OLD_UI19="documents\\UI-19.docx"
ORIGINAL_UI_28="documents\\UI-2.8.docx"
ORIGINAL_SALARY_SCHED="documents\\salary_schedule_signed.pdf"

def replaceID(id_num, document):
    id = list(str(id_num))
    while(len(id)<13):
        id.append(" ")

    document.Replace("D1", id[0], True, True)
    document.Replace("D2", id[1], True, True)
    document.Replace("D3", id[2], True, True)
    document.Replace("D4", id[3], True, True)
    document.Replace("D5", id[4], True, True)
    document.Replace("D6", id[5], True, True)
    document.Replace("D7", id[6], True, True)
    document.Replace("D8", id[7], True, True)
    document.Replace("D9", id[8], True, True)
    document.Replace("D10", id[9], True, True)
    document.Replace("D11", id[10], True, True)
    document.Replace("D12", id[11], True, True)
    document.Replace("D13", id[12], True, True)

def replaceUIF(uif_num, document):
    uif = list(str(uif_num).replace("/",""))
    while(len(uif)<7):
        uif.append(" ")

    document.Replace("U1", uif[0], True, True)
    document.Replace("U2", uif[1], True, True)
    document.Replace("U3", uif[2], True, True)
    document.Replace("U4", uif[3], True, True)
    document.Replace("U5", uif[4], True, True)
    document.Replace("U6", uif[5], True, True)
    document.Replace("U7", uif[6], True, True)
    document.Replace("U8", uif[7], True, True)

def replacePAYE(paye_num, document):
    paye = list(str(paye_num))
    while(len(paye)<10):
        paye.append(" ")

    document.Replace("P1", paye[0], True, True)
    document.Replace("P2", paye[1], True, True)
    document.Replace("P3", paye[2], True, True)
    document.Replace("P4", paye[3], True, True)
    document.Replace("P5", paye[4], True, True)
    document.Replace("P6", paye[5], True, True)
    document.Replace("P7", paye[6], True, True)
    document.Replace("P8", paye[7], True, True)
    document.Replace("P9", paye[8], True, True)
    document.Replace("P10", paye[9], True, True)

def replaceCIPC(cipc_num, document):
    cipc = list(str(cipc_num).replace("/",""))
    while(len(cipc)<16):
        cipc.append(" ")

    document.Replace("C1", cipc[0], True, True)
    document.Replace("C2", cipc[1], True, True)
    document.Replace("C3", cipc[2], True, True)
    document.Replace("C4", cipc[3], True, True)
    document.Replace("C5", cipc[4], True, True)
    document.Replace("C6", cipc[5], True, True)
    document.Replace("C7", cipc[6], True, True)
    document.Replace("C8", cipc[7], True, True)
    document.Replace("C9", cipc[8], True, True)
    document.Replace("C10", cipc[9], True, True)
    document.Replace("C11", cipc[10], True, True)
    document.Replace("C12", cipc[11], True, True)
    document.Replace("C13", cipc[12], True, True)
    document.Replace("C14", cipc[13], True, True)
    document.Replace("C15", cipc[14], True, True)
    document.Replace("C16", cipc[15], True, True)

def replaceCommencement(start_date, document):
    date = list(start_date)

    document.Replace("<<d1>>", date[0], True, True)
    document.Replace("<<d2>>", date[1], True, True)
    document.Replace("<<d3>>", date[3], True, True)
    document.Replace("<<d4>>", date[4], True, True)
    document.Replace("<<d5>>", date[8], True, True)
    document.Replace("<<d6>>", date[9], True, True)

def replaceTermination(end_date, document):
    date = list(end_date)

    document.Replace("<<t1>>", date[0], True, True)
    document.Replace("<<t2>>", date[1], True, True)
    document.Replace("<<t3>>", date[3], True, True)
    document.Replace("<<t4>>", date[4], True, True)
    document.Replace("<<t5>>", date[8], True, True)
    document.Replace("<<t6>>", date[9], True, True)

def removeWatermarks(doc_name):
    text = "Evaluation Warning: The document was created with Spire.Doc for Python."
    doc = fitz.open(doc_name+".pdf")
    for i in range(doc.page_count):
        page = doc.load_page(i)
        rl = page.search_for(text, quads=True)
        page.add_redact_annot(rl[0])
        page.apply_redactions()
        new_doc_name=doc_name+"_final.pdf"
    doc.save(new_doc_name)

def removePDFWatermarks(doc_name):
    text = "Evaluation Warning : The document was created with Spire.PDF for Python."
    doc = fitz.open(doc_name+".pdf")
    for i in range(doc.page_count):
        page = doc.load_page(i)
        rl = page.search_for(text, quads=True)
        page.add_redact_annot(rl[0])
        page.apply_redactions()
        new_doc_name=doc_name+"_final.pdf"
    doc.save(new_doc_name)

def enter_data():
        # User info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        clock_num = clock_num_entry.get()
        hrs = hrs_worked_entry.get()
        id = id_entry.get()
        reason = reason_entry.get()
        remuneration = remuneration_entry.get()
        if (str(remuneration).__contains__("R")):
            tkinter.messagebox.showwarning(title="Error", message="Please enter remuneration in the form: XXXX.cc (no R / currency)")
            return
        if not(reason or firstname or lastname or clock_num or id or remuneration or start_date or end_date or company_name or uif or paye or cipc):
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
            return
        else:
            start_date = commencement_entry.get()
            end_date = termination_entry.get()
            if len(start_date)<8 or len(end_date)<8 :
                tkinter.messagebox.showwarning(title="Error", message="Start date and end date must be in the format: DD/MM/YYYY")
                return
            else:
                company_name = company_entry.get()
                address = address_entry.get()
                uif = UIF_entry.get()
                paye = PAYE_entry.get()
                cipc = CIPC_entry.get()
                if not uif:
                    tkinter.messagebox.showwarning(title="Error", message="Please provide a UIF reference number with the format '2413510/8' to check if the company exists in the database")
                    return
                elif uif and paye and cipc and address and company_name:
                    addCompany(uif,paye,cipc,address,company_name)
                    generateUI19s(hrs,company_name, firstname, lastname, id, start_date, end_date, uif, paye, cipc, reason,
                                  address, clock_num, remuneration)

                elif not (paye or cipc or address or company_name):
                    arr = checkCompany(uif)
                    if len(arr) == 1:
                        tkinter.messagebox.showwarning(title="Error",
                                                       message="UIF number does not exist in the database")
                        return
                    else:
                        company_name = arr[0]
                        uif = arr[1]
                        paye = arr[2]
                        cipc = arr[3]
                        address = arr[4]
                        generateUI19s(hrs,company_name, firstname, lastname, id, start_date, end_date, uif, paye, cipc, reason, address, clock_num, remuneration)
                else:
                    generateUI19s(hrs,company_name, firstname, lastname, id, start_date, end_date, uif, paye, cipc, reason, address, clock_num, remuneration)

def checkCompany (uifNo):
    csv_file_path = 'C:\\Users\\FLDCLA001\\PycharmProjects\\UI19Filler\\database\\companies.csv'
    with open(csv_file_path, 'r') as file:
        csv_reader = csv.reader(file, delimiter='|')
        company_info = []
        for row in csv_reader:
            if str(row).__contains__(str(uifNo)):
                company_info.append(row)
    return company_info[0]

def addCompany(uif, paye, cipc, address, name):
    csv_file_path = 'C:\\Users\\FLDCLA001\\PycharmProjects\\UI19Filler\\database\\companies.csv'
    new_company_info = name + "|" + uif + "|" + paye + "|" + cipc + "|" + address

    # read from file
    file = open(csv_file_path, 'r')
    Lines = file.readlines()
    file.close()

    for line in Lines:
        if line.__contains__(uif):
            Lines.remove(line)

    # writing to file
    file = open(csv_file_path, 'w')
    file.writelines(Lines)
    file.write(new_company_info)
    file.write("\n")
    file.close()

def generateNewUI19(date, company, firstname, surname, id, commence, term, uif_num, paye_num, cipc_num, reason, address, clock_num, remuneration):
    # Create a Document object
    document = Document()
    # Load a Word document
    document.LoadFromFile(ORIGINAL_NEW_UI19)

    # Replace the first instance of a text with another text
    document.Replace("<<ref no>>", uif_num, True, True)
    document.Replace("<<reg no>>", cipc_num, True, True)
    document.Replace("<<paye>>", paye_num, True, True)
    document.Replace("<<company name>>", company, True, True)
    document.Replace("<<address>>", address, True, True)
    replaceID(id, document)
    replaceCommencement(commence,document)
    replaceTermination(term, document)
    calcRemuneration(document,remuneration)
    document.Replace("<<reason>>", reason, True, True)
    document.Replace("<<firstname>>", firstname, True, True)
    document.Replace("<<surname>>", surname, True, True)
    document.Replace("<<date>>", date, True, True)
    document.Replace("<<clock>>", clock_num, True, True)
    # Save the resulting document
    new_name = "C:\\Users\\FLDCLA001\\Desktop\\UI-19s\\"+company+"_"+surname+"_newUI19"
    document.SaveToFile(new_name+".pdf")
    document.Close()

    removeWatermarks(new_name)

def generateOldUI19(date, company, firstname, surname, id, commence, term, uif_num, paye_num, cipc_num, reason, address, hrs, remuneration):
    # Create a Document object
    document = Document()
    # Load a Word document
    document.LoadFromFile(ORIGINAL_OLD_UI19)

    # Replace the first instance of a text with another text
    replaceUIF(uif_num,document)
    replaceCIPC(cipc_num, document)
    replacePAYE(paye_num, document)
    document.Replace("<<company name>>", company, True, True)
    document.Replace("<<address>>", address, True, True)

    replaceID(id, document)
    replaceCommencement(commence, document)
    replaceTermination(term, document)
    calcRemuneration(document, remuneration)
    reason_list = ["", "Deceased", "Retired", "Dismissed", "Contract Expired", "Resigned", "Constructive Dismissal",
                   "Insolvency/Liquidation", "Maternity/Adoption", "Illness /Medically boarded",
                   "Retrenched/Staff Reduction", "Transfer to another Branch", "Absconded", "Business Closed",
                   "Death of Domestic Employer", "Voluntary Severance Package"]
    r_num = "2"
    for r in reason_list:
        if r.__contains__(reason):
            r_num = str((reason_list.index(r)) + 1)
    document.Replace("<<r>>", r_num, True, True)
    document.Replace("<<init>>", firstname[0], True, True)
    document.Replace("<<surname>>", surname, True, True)

    today = datetime.now()
    month = today.strftime("%b")
    document.Replace("<<date>>", date, True, True)
    document.Replace("<<MONTH>>", month.upper(), True, True)
    document.Replace("<<hrs>>", hrs, True, True)

    # Save the resulting document
    new_name = "C:\\Users\\FLDCLA001\\Desktop\\UI-19s\\" + company + "_" + surname + "_oldUI19"
    document.SaveToFile(new_name + ".pdf")
    document.Close()

    removeWatermarks(new_name)

def generateUI28(id,firstname,surname,company):
    # Create a Document object
    document = Document()
    # Load a Word document
    document.LoadFromFile(ORIGINAL_UI_28)
    replaceID(id, document)
    name = str(firstname).upper()+ " " + str(surname).upper()
    document.Replace("<<name>>",name,True,True)
    # Save the resulting document
    new_name = "C:\\Users\\FLDCLA001\\Desktop\\UI-19s\\" + company + "_" + surname + "_UI28"
    document.SaveToFile(new_name + ".pdf")
    document.Close()
    removeWatermarks(new_name)

def generateSalarySched(id, firstname, surname, uif, company, commence, term, remuneration):
    # Create an object of the PdfDocument class
    doc = PdfDocument()
    # Load a PDF file
    doc.LoadFromFile(ORIGINAL_SALARY_SCHED)

    # Iterate through the pages in the document
    for i in range(doc.Pages.Count):
        # Get the current page
        page = doc.Pages[i]
        # Create an object of the PdfTextReplace class and pass the page to the constructor of the class as a parameter
        replacer = PdfTextReplacer(page)

        # Replace All instances of a specific text with new text
        replacer.ReplaceAllText("<<<<<<<<<<<<<<<<<id number>>>>>>>>>>>>>>>>>>>", id)
        name = firstname[0] + " "+surname
        replacer.ReplaceAllText("<<<<<<<<<<<<<<<<<name>>>>>>>>>>>>>>>>>>>>>>", name)
        replacer.ReplaceAllText("<<<<<<<<<<<<<<<<<<<<ref no>>>>>>>>>>>>>>>>>>>", uif)
        replacer.ReplaceAllText("<<<<<<<<<<<<<<<<<<company name>>>>>>>>>>>>>>", company)
        replacer.ReplaceAllText("<<<<<<<<start>>>>>>>>", commence)
        replacer.ReplaceAllText("<<<<<<<<term>>>>>>>>", term)
        uif_commence = "25"+commence[2:]
        replacer.ReplaceAllText("<<<commence>>>", uif_commence)
        replacer.ReplaceAllText("<<<<<remun>>>>>", remuneration)

        contrib_amount = float(remuneration) * 0.01
        if (contrib_amount >= 1771.12):
            replacer.ReplaceAllText("<<<<<ded>>>>>", "1771.12")
        else:
            replacer.ReplaceAllText("<<<<<ded>>>>>", str(round(contrib_amount,2)))

        replacer.ReplaceAllText("<<<date>>>", str(date.today()))

    # Save the resulting document
    new_name = "C:\\Users\\FLDCLA001\\Desktop\\UI-19s\\" + company + "_" + surname + "_salarySchedule"
    doc.SaveToFile(new_name + ".pdf")
    doc.Close()
    removePDFWatermarks(new_name)

def generateUI19s(hrs,company, firstname, surname, id, commence, term, uif_num, paye_num, cipc_num, reason, address, clock_num, remuneration):
    today_date = str(date.today())
    if(not address):
        address = "1 Kingsbarns Road, Sunningdale"
    generateNewUI19(today_date,company,firstname, surname,id, commence, term, uif_num, cipc_num,paye_num,reason,address,clock_num,remuneration)
    generateOldUI19(today_date,company,firstname, surname,id, commence, term, uif_num, cipc_num,paye_num,reason,address,hrs,remuneration)
    generateUI28(id,firstname, surname, company)
    generateSalarySched(id, firstname, surname, uif_num, company, commence, term, remuneration)
    tkinter.messagebox.askokcancel("Complete","All UI-19s generated successfully")

def generateReason(reason, document):
    reason_list = ["","Deceased","Retrired","Dismissed","Contract Expired",	"Resigned", "Constructive Dismissal", "Insolvency/Liquidation", "Maternity/Adoption", "Illness /Medically boarded",	"Retrenched/Staff Reduction", "Transfer to another Branch", "Absconded", "Business Closed", "Death of Domestic Employer", "Voluntary Severance Package"]
    r_num = "2"
    for r in reason_list:
        if r.__contains__(reason):
            r_num = str((reason_list.index(r))+1)
    document.Replace("<<r>>",r_num,True,True)

def calcRemuneration(document, remuneration):
    rem_arr = remuneration.split(".")
    document.Replace("<<R1>>",rem_arr[0],True,True)
    document.Replace("<<c1>>", rem_arr[1], True, True)
    contrib_amount = float(remuneration)*0.01
    contrib_arr = str(contrib_amount).split(".")
    if(contrib_amount>=1771.12):
        document.Replace("<<R2>>", "1771", True, True)
        document.Replace("<<c2>>", "12", True, True)
    else:
        document.Replace("<<R2>>", contrib_arr[0], True, True)
        document.Replace("<<c2>>", str(contrib_arr[1])[:2], True, True)

def import_csv():
    file = fd.askopenfile()
    print(file.read())


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    #GUI--------------------------------------------------------

    window = tkinter.Tk()
    window.title("UI-19 Data Entry Form")

    frame = tkinter.Frame(window)
    frame.pack()

    # Saving Employee Info
    user_info_frame = tkinter.LabelFrame(frame, text="Employee Information")
    user_info_frame.grid(row=0, column=0, padx=30, pady=20)

    first_name_label = tkinter.Label(user_info_frame, text="First Name")
    first_name_label.grid(row=0, column=0)
    last_name_label = tkinter.Label(user_info_frame, text="Last Name")
    last_name_label.grid(row=0, column=1)

    first_name_entry = tkinter.Entry(user_info_frame)
    last_name_entry = tkinter.Entry(user_info_frame)
    first_name_entry.grid(row=1, column=0)
    last_name_entry.grid(row=1, column=1)

    clock_num_label = tkinter.Label(user_info_frame, text="Clock Number")
    clock_num_entry = tkinter.Entry(user_info_frame)
    clock_num_label.grid(row=0, column=2)
    clock_num_entry.grid(row=1, column=2)

    hrs_worked_label = tkinter.Label(user_info_frame, text="Hours Worked")
    hrs_worked_entry = tkinter.Entry(user_info_frame)
    hrs_worked_label.grid(row=2, column=2)
    hrs_worked_entry.grid(row=3, column=2)

    id_label = tkinter.Label(user_info_frame, text="ID Num")
    id_entry = tkinter.Entry(user_info_frame)
    id_label.grid(row=2, column=0)
    id_entry.grid(row=3, column=0)

    remuneration_label = tkinter.Label(user_info_frame, text="Remuneration")
    remuneration_entry = ttk.Entry(user_info_frame)
    remuneration_label.grid(row=2, column=1)
    remuneration_entry.grid(row=3, column=1)

    # Import Button
    button = tkinter.Button(frame, text="Import Employee Information", command=import_csv)
    button.grid(row=3, column=3, sticky="news", padx=20, pady=10)

    for widget in user_info_frame.winfo_children():
        widget.grid_configure(padx=10, pady=5)

    # Saving commencement & termination date
    dates_frame = tkinter.LabelFrame(frame)
    dates_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

    commencement_label = tkinter.Label(dates_frame, text="Commencement Date")
    commencement_entry = ttk.Entry(dates_frame)
    commencement_label.grid(row=0, column=1)
    commencement_entry.grid(row=1, column=1)

    termination_label = tkinter.Label(dates_frame, text="Termination Date")
    termination_entry = ttk.Entry(dates_frame)
    termination_label.grid(row=0, column=2)
    termination_entry.grid(row=1, column=2)

    reason_label = tkinter.Label(dates_frame, text="Reason for Termination")
    reason_entry = ttk.Entry(dates_frame)
    reason_label.grid(row=0, column=0)
    reason_entry.grid(row=1, column=0)

    for widget in dates_frame.winfo_children():
        widget.grid_configure(padx=10, pady=5)

    # Company Info
    company_frame = tkinter.LabelFrame(frame, text="Company Information")
    company_frame.grid(row=2, column=0,sticky="news", padx=30, pady=20)

    UIF_label = tkinter.Label(company_frame, text="UIF Number")
    UIF_entry = ttk.Entry(company_frame)
    UIF_label.grid(row=0, column=1)
    UIF_entry.grid(row=1, column=1)

    CIPC_label = tkinter.Label(company_frame, text="CIPC Number")
    CIPC_entry = ttk.Entry(company_frame)
    CIPC_label.grid(row=0, column=2)
    CIPC_entry.grid(row=1, column=2)

    PAYE_label = tkinter.Label(company_frame, text="PAYE Number")
    PAYE_entry = ttk.Entry(company_frame)
    PAYE_label.grid(row=0, column=0)
    PAYE_entry.grid(row=1, column=0)

    address_label = tkinter.Label(company_frame, text="Company Address")
    address_entry = ttk.Entry(company_frame)
    address_label.grid(row=2, column=2)
    address_entry.grid(row=3, column=2)

    company_label = tkinter.Label(company_frame, text="Company Name")
    company_entry = ttk.Entry(company_frame)
    company_label.grid(row=2, column=0)
    company_entry.grid(row=3, column=0)

    for widget in company_frame.winfo_children():
        widget.grid_configure(padx=10, pady=5)

    # Button
    button = tkinter.Button(frame, text="Generate UI19s", command=enter_data)
    button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

    window.mainloop()

    #___________________________________________________________________


