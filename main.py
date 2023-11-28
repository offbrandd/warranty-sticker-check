import csv
import tkinter as tk
import gspread
from playsound import playsound


gc = gspread.service_account(filename='keys.json')
sh = gc.open("Sprout Computer Care")

worksheet = sh.worksheet("Master List")
master_list = worksheet.get_all_values()
master_sn_col = master_list[0].index("sn")
warranty_col = master_list[0].index("warranty")

worksheet = sh.worksheet("Query result")
query_list = worksheet.get_all_values()
barcode_col = query_list[0].index("asset_barcode")
query_sn_col = query_list[0].index("device_serial")


def get_serial(barcode):
    for row in query_list:
        if row[barcode_col] == barcode:
            return row[query_sn_col]
    return "NA"

def isWarranty(sn):
    for row in master_list:
        if row[master_sn_col] == sn:
            try: 
                if row[warranty_col].index('y') > -1:
                    return True
            except ValueError:
                return False
    return False

def check_unit(barcode):
    sn = get_serial(barcode)
    if sn != "NA":
        if isWarranty(sn):
            print("Warranty confirmed")
            result_canvas.configure(bg="green")
            result_label.configure(text=barcode)
            playsound('ding.mp3', block=False)
        else:
            print("Warranty expired")
            result_canvas.configure(bg="red")
            result_label.configure(text=barcode)
    else:
        print("Could not find SN")
        result_canvas.configure(bg="yellow")
        result_label.configure(text=barcode)

def submit(self):
    barcode=barcode_var.get()
    barcode_var.set("")
    check_unit(barcode)

root = tk.Tk()
root.geometry("600x300")
root.bind('<Return>', submit)

barcode_var = tk.StringVar()
barcode_label = tk.Label(root, text = 'Enter Barcode', font=('calibre',10, 'bold'))
barcode_entry = tk.Entry(root,textvariable = barcode_var, font=('calibre',10,'normal'))
sub_btn=tk.Button(root,text = 'Submit', command = submit)
result_canvas = tk.Canvas(root, width=300, height=200, bg="yellow")
result_label = tk.Label(result_canvas, text = "pending", font=('calibre',36, 'bold'))

barcode_label.grid(row=0,column=0)
barcode_entry.grid(row=1,column=0)
sub_btn.grid(row=1,column=1)
result_canvas.grid(row=1,column=2)
result_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

if __name__ == "__main__":
    root.mainloop()

