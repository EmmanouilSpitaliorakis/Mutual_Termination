from datetime import datetime, date
import numpy as np
import tkinter as tk
from tkinter import font
from classes.Classes import Text
import os, xlsxwriter, platform 
import tkinter.messagebox as messagebox

ADMIN_FEES = 75
DAYS = 30
today = date.today()
today = today.strftime("%d%m%Y")


def useAdmin(first_payment_day, date_of_termination):
	if (date_of_termination - first_payment_day).days >= DAYS:
		return ADMIN_FEES
	else:
		return 0

def days_difference(first_date, second_date):
	return abs((first_date - second_date).days)

def transform_dates():
	first_payment_obj = datetime.strptime(first_payment_entry.get(), '%d/%m/%Y').date()
	start_date_obj = datetime.strptime(start_date_entry.get(), '%d/%m/%Y').date()
	end_date_obj = datetime.strptime(end_date_entry.get(), '%d/%m/%Y').date()
	date_of_termination_obj = datetime.strptime(date_of_termination_entry.get(), '%d/%m/%Y').date()

	return first_payment_obj, start_date_obj, end_date_obj, date_of_termination_obj

def create_folder_on_os():
	os_base = platform.system()
	if os_base != 'Darwin':
		desktop_dir = os.environ["USERPROFILE"] + '\\Desktop\\'
		if os.path.exists(desktop_dir + 'Mutual Termination\\'):
			file_dir = desktop_dir + 'Mutual Termination\\'
		else: 
			os.makedirs(desktop_dir + 'Mutual Termination\\')
			file_dir = desktop_dir + 'Mutual Termination\\'
	else:
		desktop_dir = os.environ['HOME'] + '/Desktop/'
		if os.path.exists(desktop_dir + 'Mutual Termination/'):
			file_dir = desktop_dir + 'Mutual Termination/'
		else: 
			os.makedirs(desktop_dir + 'Mutual Termination/')
			file_dir = desktop_dir + 'Mutual Termination/'
	return file_dir




def calculate_refund(first_payment_entry, fee_entry, instalments_entry, start_date_entry, end_date_entry, date_of_termination_entry):
	first_payment_entry, start_date_entry, end_date_entry, date_of_termination_entry = transform_dates()

	instalment_value = float(fee_entry / instalments_entry)
	contract_days = days_difference(start_date_entry, end_date_entry)
	cost_per_day = float(fee_entry) / contract_days
	days_used = days_difference(start_date_entry, date_of_termination_entry)
	months_used = (days_used / DAYS)
	months_paid = round((days_difference(first_payment_entry, date_of_termination_entry) / DAYS)+1.0)
	cost_used = float(cost_per_day * days_used)
	admin_cost = ADMIN_FEES
	total_cost = cost_used + admin_cost
	if instalments_entry == 1: 
		amount_paid = instalment_value
	else:
		amount_paid = months_paid * instalment_value

	outstanding_value = total_cost - amount_paid

	calculate_label['text'] = f'£ {outstanding_value}'

	return instalment_value, contract_days, cost_per_day, days_used, months_used, months_paid, cost_used, total_cost, outstanding_value

def create_excel(first_payment_entry, fee_entry, instalments_entry, start_date_entry, end_date_entry, date_of_termination_entry):
	instalment_value, contract_days, cost_per_day, days_used, months_used, months_paid, cost_used, total_cost, outstanding_value = calculate_refund(first_payment_entry, fee_entry, instalments_entry, start_date_entry, end_date_entry, date_of_termination_entry)
	file_dir = create_folder_on_os()

	if os.path.exists(file_dir + 'excels'):
		excel_dir = file_dir + 'excels/'
	else:
		os.makedirs(file_dir + 'excels')
		excel_dir = file_dir + 'excels/'

	book = xlsxwriter.Workbook(excel_dir + f'{today}_{customer_name_entry.get()}.xlsx')
	sheet = book.add_worksheet()

	format_B_row = book.add_format({'bold': True, 'font_color': '#538dd5', 'border': 3, 'fg_color': '#ffffff'})
	format_C_row = book.add_format({'align': 'center', 'border': 3, 'fg_color': '#ffffff'})
	format_C_row_bold = book.add_format({'align': 'center', 'bold': True, 'border': 3, 'fg_color': '#ffffff'})
	currency_format = book.add_format({'num_format': '£#,##0.00', 'align': 'center', 'border': 3, 'fg_color': '#ffffff'})
	sheet.set_column('B:B', 43.91)
	sheet.set_column('C:C', 22.90)

	sheet.write('B3', Text.customer_name, format_B_row)
	sheet.write('B4', Text.first_payment_day, format_B_row)
	sheet.write('B5', Text.fees, format_B_row)
	sheet.write('B6', Text.instalments, format_B_row)
	sheet.write('B7', Text.start_date, format_B_row)
	sheet.write('B8', Text.end_date, format_B_row)
	sheet.write('B9', Text.date_of_termination, format_B_row)
	sheet.write('B10', Text.contract_days, format_B_row)
	sheet.write('B11', Text.cost_per_day, format_B_row)
	sheet.write('B12', Text.days_used, format_B_row)
	sheet.write('B13', Text.months_used, format_B_row)
	sheet.write('B14', Text.months_paid, format_B_row)
	sheet.write('B15', Text.cost_used, format_B_row)
	sheet.write('B16', Text.admin_cost, format_B_row)
	sheet.write('B17', Text.total_cost, format_B_row)
	if outstanding_value < 0:
		sheet.write('B18', Text.HH_outstanding_value, format_B_row)
	else:
		sheet.write('B18', Text.customer_outstanding_value, format_B_row)


	sheet.write('C3', customer_name_entry.get(), format_C_row_bold)
	sheet.write('C4', first_payment_entry, format_C_row)
	sheet.write('C5', f'£ {fee_entry}', currency_format)
	sheet.write('C6', instalments_entry, format_C_row)
	sheet.write('C7', start_date_entry, format_C_row)
	sheet.write('C8', end_date_entry, format_C_row)
	sheet.write('C9', date_of_termination_entry, format_C_row)
	sheet.write('C10', contract_days, format_C_row)
	sheet.write('C11', f'£ {cost_per_day}', currency_format)
	sheet.write('C12', days_used, format_C_row)
	sheet.write('C13', months_used, format_C_row)
	sheet.write('C14', months_paid, format_C_row)
	sheet.write('C15', f'£ {cost_used}', currency_format)
	sheet.write('C16', f'£ {ADMIN_FEES}', currency_format)
	sheet.write('C17', f'£ {total_cost}', currency_format)
	sheet.write('C18', f'£ {abs(outstanding_value)}', currency_format)



	book.close()
	messagebox.showinfo("Export Excel", "Excel Exported.")


root = tk.Tk()
root.iconbitmap(default = r'./img/hh.ico')
root.title("Refunds Application")

WIDTH = root.winfo_screenwidth()
HEIGHT = root.winfo_screenheight()

canvas = tk.Canvas(root, height = HEIGHT * 0.50 , width = WIDTH * 0.50)
canvas.pack()

frame = tk.Frame(canvas, bd = 5, bg = '#999966')
frame.place(x = 0, y = 0, relheight = 1, relwidth = 1)


customer_name_label = tk.Label(frame, font = ('Arial', 12), text = Text.customer_name_text, bg = "#fff1cc",)
customer_name_label.place(relheight = 0.10, relwidth = 0.50)
customer_name_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#fff1cc')
customer_name_entry.place(relx = 0.52, rely = 0, relheight = 0.10, relwidth = 0.48)

first_payment_label = tk.Label(frame, font = ('Arial', 12), text = Text.first_payment_day_text, bg = "#fff1cc")
first_payment_label.place(rely = 0.12, relheight = 0.10, relwidth = 0.50)
first_payment_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#fff1cc')
first_payment_entry.place(relx = 0.52, rely = 0.12, relheight = 0.10, relwidth = 0.48)

fee_label = tk.Label(frame, font = ('Arial', 12), text = Text.fee_text, bg = "#fff1cc")
fee_label.place(rely = 0.24, relheight = 0.10, relwidth = 0.50)
fee_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#fff1cc')
fee_entry.place(relx = 0.52, rely = 0.24, relheight = 0.10, relwidth = 0.48)

instalments_label = tk.Label(frame, font = ('Arial', 12), text = Text.instalments_text, bg = "#fff1cc")
instalments_label.place(rely = 0.36, relheight = 0.10, relwidth = 0.50)
instalments_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#fff1cc')
instalments_entry.place(relx = 0.52, rely = 0.36, relheight = 0.10, relwidth = 0.48)

start_date_label = tk.Label(frame, font = ('Arial', 12), text = Text.start_date_text, bg = "#fff1cc")
start_date_label.place(rely = 0.48, relheight = 0.10, relwidth = 0.50)
start_date_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#fff1cc')
start_date_entry.place(relx = 0.52, rely = 0.48, relheight = 0.10, relwidth = 0.48)

end_date_label = tk.Label(frame, font = ('Arial', 12), text = Text.end_date_text, bg = "#fff1cc")
end_date_label.place(rely = 0.60, relheight = 0.10, relwidth = 0.50)
end_date_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#fff1cc')
end_date_entry.place(relx = 0.52, rely = 0.60, relheight = 0.10, relwidth = 0.48)

date_of_termination_label = tk.Label(frame, font = ('Arial', 12), text = Text.date_of_termination_text, bg = "#fff1cc")
date_of_termination_label.place(rely = 0.72, relheight = 0.10, relwidth = 0.50)
date_of_termination_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#fff1cc')
date_of_termination_entry.place(relx = 0.52, rely = 0.72, relheight = 0.10, relwidth = 0.48)


calculate_button = tk.Button(frame, text = "Calculate Refund", font = ('Arial', 12), bg = '#fff1cc', command = lambda: calculate_refund(first_payment_entry.get(), float(fee_entry.get()), int(instalments_entry.get()), start_date_entry.get(), end_date_entry.get(), date_of_termination_entry.get()) )
calculate_button.place(relx = 0.20, rely = 0.85,  relheight = 0.10, relwidth = 0.12)

calculate_label = tk.Label(frame, font = ('Arial', 12), bg = "#fff1cc")
calculate_label.place(relx = 0.445, rely = 0.85,  relheight = 0.10, relwidth = 0.12)

export_button = tk.Button(frame, text = "Export Excel", font = ('Arial', 12), bg = '#fff1cc', command = lambda: create_excel(first_payment_entry.get(), float(fee_entry.get()), int(instalments_entry.get()), start_date_entry.get(), end_date_entry.get(), date_of_termination_entry.get()) )
export_button.place(relx = 0.70, rely = 0.85,  relheight = 0.10, relwidth = 0.12)

root.mainloop()

