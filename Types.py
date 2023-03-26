from decimal import *


class MonthlyIncome:
    def __init__(self, hours, hours_ip, other_amount, total_amount):
        self.hours = Decimal(hours)
        self.hours_ip = Decimal(hours_ip)
        self.other_amount = Decimal(str(other_amount))
        self.amount = Decimal(str(total_amount))
        self.verify()
        return

    def verify(self):
        if self.hours < self.hours_ip:
            raise Exception("Total hours cannot be lower than hours_ip")
        if self.amount < self.other_amount:
            raise Exception("Total amount cannot be lower than other amount")

    def getIpBoxAmount(self):
        return (self.amount - self.other_amount) * self.hours_ip / self.hours

    def getIpBoxPercentage(self):
        value = self.getIpBoxAmount() / self.amount
        return value


class Invoice:
    def __init__(self, date, number, name, invoice_type, amount):
        self.date = date
        self.number = str(number)
        self.name = str(name)
        self.type = str(invoice_type)
        self.amount = Decimal(str(amount))


InvoiceList = list[Invoice]
IncomeData = dict[int, MonthlyIncome]
InvoicesGrouped = dict[int, InvoiceList]
IpAllocation = dict[int, int]


class InputData:
    nexus_invoices: InvoicesGrouped
    cost_invoices: InvoicesGrouped
    income_invoices: InvoicesGrouped
    income_data: IncomeData
    ip_allocation: IpAllocation
    other: dict[str, str]
