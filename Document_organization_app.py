import tkinter as tk
from os import rename, walk, path
from PyPDF2 import PdfFileMerger, PdfFileReader
from datetime import datetime
import xlwings as xw
import pdfkit

name = "Your_Name" #make sure you use "_" between your first, middle and last name (i.e.: first_middle_last)
def document_org(): 
    #get the website in pdf format 

    #get variables
    language = language_in.get()    #which language is CV or CL (currently only eng/other)
    job = str(job_in.get()).upper() #company name
    cl = cl_in.get()                #optional cover letter
    location = location_in.get()    #where is the job stationed
    position = position_in.get()    #what kind of position are you applying to
    url = str(url_in.get())         #url from je job listing
    dir="where_you_wish_to_store"   #direcotry where you wish to store merged PDF

    path_wkthmltopdf = "path_to_installed_wkhtmltopdf/bin/wkhtmltopdf.exe"
    config = pdfkit.configuration(wkhtmltopdf = path_wkthmltopdf)
    pdfkit.from_url(url, dir + str(job) + ".pdf", configuration=config)
    Current_Date = Current_Date = datetime.now().strftime('%Y-%d-%m')

    ########CV######
    if language == "eng": 
        folder_in = "path_to_stored_cv.pdf"
        folder_out = "path_where_you_wish_to_store_documents" 
        rename(folder_in,folder_out +'CV_%s_%s_%s.pdf' %(name, Current_Date, job))       
    if language == "other":  #switch from other to any other language so if you please. 
        folder_in = "path_to_stored_cv.pdf"
        folder_out = "path_where_you_wish_to_store_documents"       
        rename(folder_in,folder_out + 'CV_%s_%s_%s.pdf' % (name, Current_Date, job))
    #####CL#####     
    if cl == "y":   
        if language == "eng": 
            folder_cl_in = "path_to_stored_cv.pdf"
            folder_cl_out = "path_to_where_you_wish_to_store" 
            rename(folder_cl_in,folder_cl_out + 'CL_%s_%s_%s.pdf' %(name, Current_Date, job))            
        if language == "other": 
            folder_cl_in = "path_to_stored.pdf"
            folder_cl_out = "path_whre_you_wish_to_store"   
            rename(folder_cl_in,folder_cl_out + 'CL_%s_%s_%s.pdf' %(name, Current_Date, job))
    else: 
        pass
    new_CV['text'] = "New name of the file: CV_%s_%s_%s.pdf" % (name, Current_Date, job)
    if cl == "y":
        new_CL['text'] = ("New name of the file: CL_%s_%s_%s.pdf" % (name, Current_Date, job))
    else: 
        pass
    
    #merge all files that are important to a job application (order: job website, CL, CV)
    merger = PdfFileMerger()
    for root,dirs,files in walk(dir):
        for filename in files: 
            if filename.endswith('.pdf') and str(job) in filename:
                  print(filename)
                  filepath = path.join(root, filename)
                  #print(filepath)
                  merger.append(PdfFileReader(open(filepath, 'rb')), import_bookmarks=False)
    merger.write(dir + str(lokacija) +'\\' + str(Current_Date) + '_' + (str(job) + '.pdf')) 
    merg_pdf['text'] = ("New name of the merged pdf: %s_%s.pdf" %(Current_Date, job))
    
    #write to Excel file
    Li = [job, position, location, Current_Date]
    xw.App().visible = False
    wb = xw.Book("path/name_of_the_xlsx")  
    Sheet1 = wb.sheets[0]
    last_cell_value = Sheet1.range('A' + str(Sheet1.cells.last_cell.row)).end('up').row
    c = 1
    for i in Li:
        Sheet1.range(last_cell_value+1, c).value = str(i) 
        c+=1
    wb.save()
    wb.close()
    
    
    final['text'] = ("Your Documents are now organized!")
    return

##########SET UP  GUI FOR APPLICATION#################
window = tk.Tk()
window.title("Job Application File Creator")
window.resizable(width=False, height=False)
#set entry points
frm_entry = tk.Frame(master=window)
job_in = tk.Entry(master=frm_entry, width=50)
language_in = tk.Entry(master=frm_entry, width=50)
cl_in = tk.Entry(master=frm_entry, width=50)
location_in = tk.Entry(master=frm_entry, width=50)
position_in = tk.Entry(master=frm_entry, width=50)
url_in= tk.Entry(master=frm_entry, width=50)


ent_language = tk.Label(master=frm_entry, text=r"Language")
ent_job = tk.Label(master=frm_entry, text=r"Company Name")
ent_cl = tk.Label(master=frm_entry, text=r"You need a cover letter? (y/n)?")
ent_location = tk.Label(master=frm_entry, text=r"Location")
ent_position =  tk.Label(master=frm_entry, text=r"Position")
ent_url = tk.Label(master=frm_entry, text=r"Url")

####set grid positions and GUI

language_in.grid(row=0, column=1, sticky="e")
job_in.grid(row=1, column=1, sticky="e")
cl_in.grid(row=2, column=1, sticky="e")
location_in.grid(row=3, column=1, sticky="e")
position_in.grid(row=4, column=1, sticky="e")
url_in.grid(row=5, column=1, sticky="e")


ent_language.grid(row=0, column=0, sticky="w")
ent_job.grid(row=1, column=0, sticky="w")
ent_cl.grid(row=2, column=0, sticky="w")
ent_location.grid(row=3, column=0, sticky="w")
ent_position.grid(row=4, column=0, sticky="w")
ent_url.grid(row=5, column=0, sticky="w")



# # # Create the conversion Button and result display Label
btn_convert = tk.Button(
    master=window,
    text=r"Submit",
    command=document_org)

new_CV = tk.Label(master=window, text=r"CV Name: ")
new_CL = tk.Label(master=window, text="CL Name: ")
merg_pdf = tk.Label(master=window, text="Merged PDF Name: ")
final = tk.Label(master=window, text="Status: ")
# Set-up the layout using the .grid() geometry manager
frm_entry.grid(row=0, column=0, padx=20)
btn_convert.grid(row=1, column=0, pady=20)
new_CV.grid(row=2, column=0, padx=10)
new_CL.grid(row=3, column=0, padx=10)
merg_pdf.grid(row=4, column=0, padx=10)
final.grid(row=5, column=0, padx=10)

# Run the application
window.mainloop()
