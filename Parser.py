import pandas as pd
import Types
from collections import defaultdict as ddict


def readIncomeData():
    income_data = pd.read_excel('input.xlsx', sheet_name=1)
    income_parsed = {}
    for row in income_data.iterrows():
        row_data = row[1]
        month = row_data['Miesiąc']
        if month in range(1, 13):
            if month in income_parsed.keys():
                raise Exception("Expected only one month: month number " + str(month))
            income_parsed[month] = Types.MonthlyIncome(row_data['Godziny'],
                                                       row_data['Godziny IP'],
                                                       row_data['Inne'],
                                                       row_data['Kwota na fakturze'])
    return income_parsed


def makeInvoice(row_data, amount):
    return Types.Invoice(row_data['Data'],
                         row_data['NR'],
                         row_data['Nazwa'],
                         row_data['Typ'],
                         amount)


def readInvoiceData() -> tuple[Types.InvoicesGrouped, Types.InvoicesGrouped, Types.InvoicesGrouped]:
    cost_invoices = ddict(list)
    nexus_invoices = ddict(list)
    income_invoices = ddict(list)
    invoice_data = pd.read_excel('input.xlsx', sheet_name=0)

    for row in invoice_data.iterrows():
        row_data = row[1]
        month = row_data['Miesiac']
        if month in range(1, 13):
            if row_data['Typ'] == 'sprzedaż':
                income_invoices[month].append(makeInvoice(row_data, row_data['Przychód']))
            else:
                invoice = makeInvoice(row_data, row_data['Wydatek'])
                if row_data['Nexus'] == 'A':
                    nexus_invoices[month].append(invoice)
                else:
                    cost_invoices[month].append(invoice)

    return cost_invoices, nexus_invoices, income_invoices


def readIpAllocation():
    ip_alloc = pd.read_excel('input.xlsx', sheet_name=2)
    ret_val = {}
    for row in ip_alloc.iterrows():
        row_data = row[1]
        month = row_data['Miesiąc']
        if month in range(1, 13):
            ret_val[month] = row_data['Nr IP']
    return ret_val

def readOthers():
    others = pd.read_excel('input.xlsx', sheet_name=3)
    return dict(zip(others.iloc[:, 0], others.iloc[:, 1]))


