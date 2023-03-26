from decimal import *

import Types


def sumAllMonths(invoices_grouped_by_month):
    sum = Decimal()
    for single_month in invoices_grouped_by_month.values():
        sum += sumInvoices(single_month)
    return sum

def sumInvoices(invoices):
    sum = Decimal()
    if isinstance(invoices, dict):
        iterable = invoices.values()
    else:
        iterable = invoices
    for j in iterable:
        sum += j.amount
    return sum

def verifyIncomeData(income_data, income_invoices):
    if sumAllMonths(income_invoices) != sumInvoices(income_data):
        raise Exception("Difference in income_data vs income invoices, please check your input")

def printInput(input_data: Types.InputData):
    sum_of_nexus = sumAllMonths(input_data.nexus_invoices)
    sum_of_costs = sum_of_nexus + sumAllMonths(input_data.cost_invoices)
    print("Please verify if totals match for input data!")
    print("Sum of all cost invoices: " + str(sum_of_costs))
    print("Nexus costs: " + str(sum_of_nexus))
    print("Total income: " + str(sumAllMonths(input_data.income_invoices)))
    print("Number of different IP: " + str(len(set(input_data.ip_allocation.values()))))
    print(input_data.other)
