import tkinter as tk
import button_openpyxl as bm
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory

def go():
            
    # Design the interface
    window = tk.Tk()
    window.title("Highlight Cells")
    window.rowconfigure(0, minsize=80, weight=1)
    window.columnconfigure(1, minsize=80, weight=1)

    fr_buttons_1 = tk.Frame(window, relief=tk.RAISED, bd=2)
    fr_labels_1  = tk.Frame(window, relief=tk.RAISED, bd=2)

    #Set labels
    lbl_limit    = tk.Label(master = fr_labels_1, text = "Cohort_size limit")
    lbl_Guid_1   = tk.Label(master = fr_labels_1, text = "Table(s) to be processed:")
    lbl_Result_1 = tk.Label(master = fr_labels_1, text = "Null")
    lbl_Status_info = tk.Label(master = fr_labels_1, text = "Status:")
    lbl_Status_curr = tk.Label(master = fr_labels_1, text = "Not start")

    # Set entrys 
    ent_limit = tk.Entry(master=fr_labels_1, width=10)

    # Set buttons 
    btn_open_xlsx = tk.Button(fr_buttons_1, text="Open file",   command = lambda:bm.openSingleXlsx(lbl_Result_1))
#    btn_open_fol = tk.Button(fr_buttons_1, text="Open folder", command = lambda:bm.openFolder(lbl_Result_1))

    btn_process = tk.Button(fr_buttons_1, text="Process and save", command = lambda:bm.processing(lbl_Status_curr,ent_limit.get()))
    #btn_confirm = tk.Button(fr_labels_1, text="confirm", command = get_entries)
       
    
    # Attach widgets to frames 
    btn_open_xlsx.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
#    btn_open_fol.grid(row=0, column=1, sticky="ew", padx=5)
    btn_process.grid(row=2, column=0, sticky="ew", padx=5)
    #btn_confirm.grid(row=1, column=2, sticky="ew", padx=5)
    
    lbl_Guid_1.grid(row=0, column=0, sticky="ew")
    lbl_Result_1.grid(row=0, column=1, sticky="ew")

    lbl_limit.grid(row=1, column=0, sticky="ew")
    ent_limit.grid(row=1, column=1, sticky="ew")

    lbl_Status_info.grid(row=3, column=0, sticky="ew")
    lbl_Status_curr.grid(row=3, column=1, sticky="ew")

    # Attach frams to windows 
    fr_buttons_1.grid(row=0, column=0, sticky="nsew")
    fr_labels_1.grid(row=0, column=1, sticky="nsew")

    window.mainloop()


