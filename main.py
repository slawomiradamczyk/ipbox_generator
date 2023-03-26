import os
import time
import Parser
import Types
import Utils
import XlsxGenerator
from decimal import *

if __name__ == '__main__':
    input_data: Types.InputData = Types.InputData()
    getcontext().prec = 28
    input_data.income_data = Parser.readIncomeData()
    input_data.cost_invoices, input_data.nexus_invoices, input_data.income_invoices = Parser.readInvoiceData()
    input_data.ip_allocation = Parser.readIpAllocation()
    input_data.other = Parser.readOthers()

    Utils.verifyIncomeData(input_data.income_data, input_data.income_invoices)
    Utils.printInput(input_data)

    generator = XlsxGenerator.Generator(os.path.abspath("result.xlsx"), input_data)
    with generator:
        generator.generateMainSheet()
        generator.generateSideSheet()
