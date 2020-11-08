from datetime import datetime
import numpy as np
import tkinter as tk
from tkinter import font
from classes.Classes import Text

ADMIN_FEES = 75
DAYS = 30

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




def calculate_refund(first_payment_entry, fee_entry, instalments_entry, start_date_entry, end_date_entry, date_of_termination_entry):
	first_payment_entry, start_date_entry, end_date_entry, date_of_termination_entry = transform_dates()
	instalment_value = round(float(fee_entry / instalments_entry), 2)

	contract_days = days_difference(start_date_entry, end_date_entry)

	cost_per_day = round(float(fee_entry) / contract_days, 2)

	days_used = days_difference(start_date_entry, date_of_termination_entry)

	months_used = (round(days_used / DAYS, 2) + 1)
	
	cost_used = float(cost_per_day * days_used)

	admin_cost = ADMIN_FEES
	total_cost = cost_used + admin_cost

	if instalments_entry == 1: 
		amount_paid = instalment_value
	else:
		amount_paid = months_used * instalment_value

	outstanding_value = round(total_cost - amount_paid, 2)

	calculate_label['text'] = f'£ {outstanding_value}'

def create_excel(first_payment_entry, fee_entry, instalments_entry, start_date_entry, end_date_entry, date_of_termination_entry):
	calculate_refund(first_payment_entry, fee_entry, instalments_entry, start_date_entry, end_date_entry, date_of_termination_entry)

root = tk.Tk()

WIDTH = root.winfo_screenwidth()
HEIGHT = root.winfo_screenheight()

canvas = tk.Canvas(root, height = HEIGHT * 0.50 , width = WIDTH * 0.50)
canvas.pack()

frame = tk.Frame(canvas, bd = 5, bg = '#2A9AA9')
frame.place(x = 0, y = 0, relheight = 1, relwidth = 1)


customer_name_label = tk.Label(frame, font = ('Arial', 12), text = Text.customer_name_text, bg = "#ffe299",)
customer_name_label.place(relheight = 0.10, relwidth = 0.50)
customer_name_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#ffe299')
customer_name_entry.place(relx = 0.52, rely = 0, relheight = 0.10, relwidth = 0.48)

first_payment_label = tk.Label(frame, font = ('Arial', 12), text = Text.first_payment_day_text, bg = "#ffe299")
first_payment_label.place(rely = 0.12, relheight = 0.10, relwidth = 0.50)
first_payment_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#ffe299')
first_payment_entry.place(relx = 0.52, rely = 0.12, relheight = 0.10, relwidth = 0.48)

fee_label = tk.Label(frame, font = ('Arial', 12), text = Text.fee_text, bg = "#ffe299")
fee_label.place(rely = 0.24, relheight = 0.10, relwidth = 0.50)
fee_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#ffe299')
fee_entry.place(relx = 0.52, rely = 0.24, relheight = 0.10, relwidth = 0.48)

instalments_label = tk.Label(frame, font = ('Arial', 12), text = Text.instalments_text, bg = "#ffe299")
instalments_label.place(rely = 0.36, relheight = 0.10, relwidth = 0.50)
instalments_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#ffe299')
instalments_entry.place(relx = 0.52, rely = 0.36, relheight = 0.10, relwidth = 0.48)

start_date_label = tk.Label(frame, font = ('Arial', 12), text = Text.start_date_text, bg = "#ffe299")
start_date_label.place(rely = 0.48, relheight = 0.10, relwidth = 0.50)
start_date_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#ffe299')
start_date_entry.place(relx = 0.52, rely = 0.48, relheight = 0.10, relwidth = 0.48)

end_date_label = tk.Label(frame, font = ('Arial', 12), text = Text.end_date_text, bg = "#ffe299")
end_date_label.place(rely = 0.60, relheight = 0.10, relwidth = 0.50)
end_date_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#ffe299')
end_date_entry.place(relx = 0.52, rely = 0.60, relheight = 0.10, relwidth = 0.48)

date_of_termination_label = tk.Label(frame, font = ('Arial', 12), text = Text.date_of_termination_text, bg = "#ffe299")
date_of_termination_label.place(rely = 0.72, relheight = 0.10, relwidth = 0.50)
date_of_termination_entry = tk.Entry(frame, font = ('Arial', 12), bg = '#ffe299')
date_of_termination_entry.place(relx = 0.52, rely = 0.72, relheight = 0.10, relwidth = 0.48)


calculate_button = tk.Button(frame, text = "Calculate Refund", font = ('Arial', 12), bg = 'white', command = lambda: calculate_refund(first_payment_entry.get(), float(fee_entry.get()), int(instalments_entry.get()), start_date_entry.get(), end_date_entry.get(), date_of_termination_entry.get()) )
calculate_button.place(relx = 0.20, rely = 0.85,  relheight = 0.10, relwidth = 0.12)

calculate_label = tk.Label(frame, font = ('Arial', 12), bg = "#ffe299")
calculate_label.place(relx = 0.445, rely = 0.85,  relheight = 0.10, relwidth = 0.12)

export_button = tk.Button(frame, text = "Export Excel", font = ('Arial', 12), bg = 'white', command = lambda: create_excel(first_payment_entry.get(), float(fee_entry.get()), int(instalments_entry.get()), start_date_entry.get(), end_date_entry.get(), date_of_termination_entry.get()) )
export_button.place(relx = 0.70, rely = 0.85,  relheight = 0.10, relwidth = 0.12)

root.mainloop()

