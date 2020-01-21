from openpyxl.worksheet.worksheet import Worksheet

from Erp import ERP
from Excel import Excel


if __name__ == "__main__":
    excel: Excel = Excel()
    excel.print_status()
    company_selected = excel.input_company()
    date_GL = excel.input_date_GL()
    excel.save_data(company_selected, date_GL)
    sheet = excel.get_selected_sheet(company_selected)

    erp: ERP = ERP()
    erp.login()
    erp.go_to_P0411()
    erp.click_add()
    erp.set_P0411_value(sheet)
    erp.set_next_P0411_value(sheet)
