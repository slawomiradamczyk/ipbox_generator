import os
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

import Types
import Utils


class Generator:
    input: Types.InputData
    number_of_ip: int
    workbook: xlsxwriter.Workbook

    side_costs_sheet: xlsxwriter.worksheet
    main_sheet: xlsxwriter.worksheet
    current_sheet: xlsxwriter.worksheet

    next_col: int
    next_row: int

    main_table_length: int

    def __init__(self, filename: os.path, input_data: Types.InputData):
        self.next_col = 0
        self.next_row = 0
        self.input = input_data
        self.number_of_ip = len(set(input_data.ip_allocation.values()))
        self.kpwi_list = ["KPWI {:d}/{:d}".format(i + 1, self.input.other['Rok']) for i in range(self.number_of_ip)]
        self.filename = filename

        self.main_table_length = 12 + (3 * self.number_of_ip)
        self.side_table_length = 7 + 1 * self.number_of_ip

    def __enter__(self):
        self.workbook = xlsxwriter.Workbook(self.filename)
        self.main_sheet = self.workbook.add_worksheet(name="Ewidencja Główna")
        self.side_costs_sheet = self.workbook.add_worksheet(name="Koszty poboczne")

    def __exit__(self, *args):
        self.workbook.close()

    def generateMainSheet(self):
        self.current_sheet = self.main_sheet
        self.generateMainTable()
        self.generateNexusCostTable()
        self.generateIncomeTable()

    def generateMainTable(self):
        self.generateMainHeader()
        self.generateMonthSummary()

    def write_groups(self, titles):
        for title, length in titles:
            if length > 1:
                self.current_sheet.merge_range(self.next_row, self.next_col, self.next_row, self.next_col + length - 1,
                                           title)
            else:
                self.current_sheet.write_string(self.next_row, self.next_col, title)
            self.next_col += length

    def writeGroupedTitles(self):
        titles = [("Czas świadczenia usług", 3),
                  ("Czas poświęcony na prace B+R przypadający na każde KPWI", 1 + self.number_of_ip),
                  ("Dochody (PLN)", 4),
                  ("Dochody przypadające na poszczególne KPWI (PLN)", 1 + self.number_of_ip),
                  ("Koszty (PLN)", 2),
                  ("Koszty przypadające na poszególne KPWI (PLN)", self.number_of_ip)]
        self.next_col = 1
        self.write_groups(titles)
        self.advance_row()

    def write_titles(self, titles):
        for title in titles:
            self.current_sheet.write_string(self.next_row, self.next_col, title)
            self.next_col += 1

    def advance_row(self, number_of_rows=1):
        self.next_col = 0
        self.next_row += number_of_rows

    def writeSmallTitles(self):
        titles = ["Prace B+R (liczba godzin)",
                  "Usługi niezaliczane do B+R (liczba godzin)",
                  "Suma",
                  ]
        titles.extend(self.kpwi_list)
        titles.extend(["Suma",
                       "Dochody netto ogółem (PLN)",
                       "Część dochodu poza rozliczeniem godzinowym",
                       "Dochody z KPWI uwzględnione w cenie usługi",
                       "Pozostała część dochodu godzinowego (niezwiązana z B+R)"])
        titles.extend(self.kpwi_list)
        titles.extend(["Suma",
                       "Koszty zaliczane zarówno do B+R i  pozostałej działalności",
                       "Koszty zaliczane zarówno do B+R i  pozostałej działalności - według proporcji przypadające na B+R"])
        titles.extend(self.kpwi_list)

        self.next_col = 1
        self.write_titles(titles)
        self.advance_row()

    def generateMainHeader(self):

        self.current_sheet.merge_range(0, 0, 0, self.main_table_length - 1,
                                       "Monitorowanie i śledzenie efektów prac badawczo-rozwojowych (B+R) w %d roku" % self.input.other['Rok'])

        self.current_sheet.merge_range(1, 0, 2, self.main_table_length - 1,
                                       "Niniejsze zestawienie ma na celu zaprezentowanie zestawienia dochodów oraz kosztów związanych z komercjalizacją kwalifikowanego prawa własności intelektualnej dla przedsiębiorstwa " + self.input.other['Dane firmy'], )
        self.current_sheet.merge_range(3, 0, 4, 0, "Lp.")
        self.advance_row(3)
        self.writeGroupedTitles()
        self.writeSmallTitles()

    def generate_invoice_list(self, name: str, invoices: Types.InvoicesGrouped):
        self.current_sheet.merge_range(self.next_row, self.next_col, self.next_row, self.next_col + 4, name)
        self.advance_row()
        titles = ['Data', "Nr Faktury", "Nazwa kosztu", "Kwota netto"]
        self.write_titles(titles)
        self.advance_row()
        total_count = 0
        for invoices_list in invoices.values():
            for invoice in invoices_list:
                self.write_and_move_col(invoice.date, self.current_sheet.write_datetime)
                self.write_and_move_col(invoice.number, self.current_sheet.write_string)
                self.write_and_move_col(invoice.type, self.current_sheet.write_string)
                self.write_and_move_col(invoice.amount, self.current_sheet.write_number)
                total_count += 1
                self.advance_row()
        self.current_sheet.merge_range(self.next_row, self.next_col, self.next_row, self.next_col + 2, "Suma:")
        self.next_col += 3
        self.write_sum_formula_vert(total_count)
        self.advance_row()

    def generateNexusCostTable(self):
        self.generate_invoice_list("Koszta Nexus", self.input.nexus_invoices)
        self.advance_row()

    def generateIncomeTable(self):
        self.generate_invoice_list("Przychody", self.input.income_invoices)
        self.advance_row()

    def generateMonthSummary(self):
        for i in self.input.income_data:
            self.generateSingleMonth(i)
            self.advance_row()
        self.write_and_move_col("Suma:", self.current_sheet.write_string)
        for i in range(self.main_table_length - 1):
            self.write_sum_formula_vert(12)
        self.advance_row(2)

    def write_sum_formula_vert(self, num_rows, suffix=''):
        # get the cell address in A1 notation
        cell = xl_rowcol_to_cell(self.next_row, self.next_col)
        # get the address of the first cell in the range to sum
        first = xl_rowcol_to_cell(self.next_row - num_rows, self.next_col)
        # get the address of the last cell in the range to sum
        last = xl_rowcol_to_cell(self.next_row - 1, self.next_col)
        # write a formula to sum the range of cells
        formula = '=SUM(%s:%s)%s' % (first, last, suffix)
        self.current_sheet.write_formula(cell, formula)
        self.next_col += 1

    def write_sum_formula(self, num_cols):
        # get the cell address in A1 notation
        cell = xl_rowcol_to_cell(self.next_row, self.next_col)
        # get the address of the first cell in the range to sum
        first = xl_rowcol_to_cell(self.next_row, self.next_col - num_cols)
        # get the address of the last cell in the range to sum
        last = xl_rowcol_to_cell(self.next_row, self.next_col - 1)
        # write a formula to sum the range of cells
        formula = '=SUM(%s:%s)' % (first, last)
        self.current_sheet.write_formula(cell, formula)
        self.next_col += 1

    def write_ip_section(self, month, value):
        ip_number = self.input.ip_allocation[month]
        for i in range(self.number_of_ip):
            self.current_sheet.write_number(self.next_row, self.next_col + i, 0)
        self.current_sheet.write_number(self.next_row, self.next_col + ip_number - 1, value)
        self.next_col += self.number_of_ip
        self.write_sum_formula(self.number_of_ip)

    def write_and_move_col(self, value, fun):
        fun(self.next_row, self.next_col, value)
        self.next_col += 1

    def write_income_formulas(self):
        # Convert row and col to cell notation
        # Convert row and col to previous columns
        income_tot_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, self.next_col - 2)
        income_other_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, self.next_col - 1)
        hours_total_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, 3)
        hours_non_kpi_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, 2)
        hours_kpi_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, 1)
        formula_kpi = '=(%s-%s)*%s/%s' % (income_tot_f, income_other_f, hours_kpi_f, hours_total_f)
        self.current_sheet.write_formula(self.next_row, self.next_col, formula_kpi)
        formula_rest = '=(%s-%s)*%s/%s' % (income_tot_f, income_other_f, hours_non_kpi_f, hours_total_f)
        self.current_sheet.write_formula(self.next_row, self.next_col + 1, formula_rest)
        self.next_col += 2

    def generateSingleMonth(self, i: int):
        income_data = self.input.income_data[i]
        self.write_and_move_col(i, self.current_sheet.write_number)
        self.write_and_move_col(income_data.hours_ip, self.current_sheet.write_number)
        self.write_and_move_col(income_data.hours - income_data.hours_ip, self.current_sheet.write_number)
        self.write_sum_formula(2)
        self.write_ip_section(i, income_data.hours_ip)
        self.write_and_move_col(income_data.amount, self.current_sheet.write_number)
        self.write_and_move_col(income_data.other_amount, self.current_sheet.write_number)
        self.write_income_formulas()
        self.write_income_per_kwpi()
        self.write_costs(self.input.nexus_invoices[i])
        self.write_costs_per_kpwi()

    def write_income_per_kwpi(self):
        income_tot_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, self.next_col - 2)
        hours_total_kpi_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, 1)
        for i in range(self.number_of_ip):
            hours_ip_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, 4 + i)
            formula_kpi = '=(%s/%s)*%s' % (hours_ip_f, hours_total_kpi_f, income_tot_f)
            self.write_and_move_col(formula_kpi, self.current_sheet.write_formula)
        self.write_sum_formula(self.number_of_ip)

    def write_costs(self, nexus_invoices: Types.InvoiceList):

        sum_of_nexus = Utils.sumInvoices(nexus_invoices)
        self.write_and_move_col(sum_of_nexus, self.current_sheet.write_number)
        income_kpi_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, self.next_col - 4 - self.number_of_ip)
        income_total_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, self.next_col - 6 - self.number_of_ip)
        nexus_total_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, self.next_col - 1)
        nexus_formula = '%s/%s * %s' % (income_kpi_f, income_total_f, nexus_total_f)
        self.write_and_move_col(nexus_formula, self.current_sheet.write_formula)
        pass

    def write_costs_per_kpwi(self):
        costs_tot_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, self.next_col - 1)
        hours_total_kpi_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, 1)
        for i in range(self.number_of_ip):
            hours_ip_f = xlsxwriter.utility.xl_rowcol_to_cell(self.next_row, 4 + i)
            formula_kpi = '=(%s/%s)*%s' % (hours_ip_f, hours_total_kpi_f, costs_tot_f)
            self.write_and_move_col(formula_kpi, self.current_sheet.write_formula)

    def print_costs_header(self, month_number):
        self.current_sheet.merge_range(self.next_row, self.next_col, self.next_row,
                                       self.next_col + self.side_table_length - 1,
                                       'Zestawienie kosztów pośrednich (tzn. niezaliczanych do wskażnika nexus) za miesiąc nr: %d roku %d' % (
                                           month_number, self.input.other['Rok']))
        self.advance_row()
        self.current_sheet.merge_range(self.next_row, self.next_col, self.next_row + 1, self.next_col, 'Lp.')
        titles = [('Dokument księgowy', 2),
                  ('Kontrahent', 1),
                  ('Opis zdarzenia', 1),
                  ('Koszt uzyskania przychodu', 1),
                  ('Część kosztu przypisana w proporcji do danego KPWI', self.number_of_ip+1)
                  ]
        self.next_col = 1
        self.write_groups(titles)
        self.advance_row()
        self.next_col = 1
        titles2 = ['Data', 'Numer', 'Nazwa', 'Czego dotyczy koszt?', 'Kwota netto']
        titles2.extend(self.kpwi_list)
        titles2.append('Suma')
        self.write_titles(titles2)


    def print_cost_invoices(self, month_number):
        cost_invoices = self.input.cost_invoices[month_number]
        income_data = self.input.income_data[month_number]
        count = 1
        for invoice in cost_invoices:
            self.write_and_move_col(count, self.current_sheet.write_number)
            self.write_and_move_col(invoice.date, self.current_sheet.write_datetime)
            self.write_and_move_col(invoice.number, self.current_sheet.write_string)
            self.write_and_move_col(invoice.name, self.current_sheet.write_string)
            self.write_and_move_col(invoice.type, self.current_sheet.write_string)
            self.write_and_move_col(invoice.amount, self.current_sheet.write_number)
            cost_part = invoice.amount * income_data.getIpBoxPercentage()
            self.write_ip_section(month_number, cost_part)
            self.advance_row()
            count += 1

    def print_costs_for_one_month(self, month_number):
        self.print_costs_header(month_number)
        self.advance_row()
        self.print_cost_invoices(month_number)
        self.current_sheet.merge_range(self.next_row, self.next_col, self.next_row, self.next_col+4, "Suma za dany miesiąc:")
        self.next_col += 4
        for i in range(6):
            self.write_sum_formula_vert(len(self.input.cost_invoices[month_number]))
        self.advance_row(2)

    def generateSideSheet(self):
        self.current_sheet = self.side_costs_sheet
        self.next_row = self.next_col = 0
        self.print_sidecost_header()
        for i in range(12):
            self.print_costs_for_one_month(i+1)
            pass

        self.current_sheet.merge_range(self.next_row, self.next_col, self.next_row, self.next_col + 4, "Suma za cały rok:")
        self.next_col += 4
        for i in range(6):
            self.write_sum_formula_vert(self.next_row, '/2')


    def print_sidecost_header(self):
        self.current_sheet.merge_range(0, 0, 0, self.side_table_length - 1,
                                       "Ewidencja kosztów pośrednich związanych z wytwarzanymi KPWI")
        self.current_sheet.merge_range(1, 0, 1, self.side_table_length - 1, "Sporządzono dla: %s" % self.input.other['Dane firmy'])
        self.advance_row(3)
        pass