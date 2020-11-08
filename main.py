from datetime import datetime
import numpy as np

ADMIN_FEES = 75
DAYS = 30

def useAdmin(first_payment_day, date_of_termination):
	if (date_of_termination - first_payment_day).days >= DAYS:
		return ADMIN_FEES
	else:
		return 0

def days_difference(first_date, second_date):
	return abs((first_date - second_date).days)

customer_name = str(input('Please enter the customer name: '))
first_payment_day = datetime.strptime(input('Please enter the date the customer paid us for the first time(dd/mm/yyyy): '), '%d/%m/%Y').date()
fee = float(input('Please enter the Fee as per Contract: '))
instalments = np.int8(input('Please enter the amount of instalments: '))
instalment_value = float(fee / instalments)
print(f'Instalment Value: {instalment_value}')
start_date = datetime.strptime(input('Please enter the tenancy start date(dd/mm/yyyy): '), '%d/%m/%Y').date()
end_date = datetime.strptime(input('Please enter the tenancy end date(dd/mm/yyyy): '), '%d/%m/%Y').date()
contract_days = days_difference(start_date, end_date)
print(f'Contract Days: {contract_days}')
cost_per_day = fee / contract_days
print(f'Cost per Day: {cost_per_day}')
date_of_termination = datetime.strptime(input('Please enter the date of mutual termination(dd/mm/yyyy): '), '%d/%m/%Y').date()
days_used = days_difference(start_date, date_of_termination)
print(f'Days Used: {days_used}')
months_used = (round(days_used / DAYS) + 1)
print(f'Months Used: {months_used}')
cost_used = float(cost_per_day * days_used)
admin_cost = useAdmin(first_payment_day, date_of_termination)
total_cost = cost_used + admin_cost
print(f'Total Cost: {total_cost}')
amount_paid = months_used * instalment_value
print(f'Amount Paid: {amount_paid}')
outstanding_value = total_cost - amount_paid
if outstanding_value > 0:
	print(f'The customer need to pay us: {outstanding_value}')
else:
	print(f'We need to pay the customer: {outstanding_value}')