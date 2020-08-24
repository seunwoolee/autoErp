import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class Excel:
    def __init__(self):
        self.wb: Workbook = openpyxl.load_workbook('authErp.xlsx')

    def save_data(self, company_selected: str, date_GL: str) -> None:
        sheet1: Worksheet = self.wb.active
        company_selected = int(company_selected) + 1
        sheet1[f'C{company_selected}'] = date_GL
        self.wb.save('authErp.xlsx')

    def print_status(self) -> None:
        rows = self.wb['Sheet1']['A2':'D15']
        for i, row in enumerate(rows):
            print(f'{i + 1}번 {row[0].value}({row[1].value}) 금액:{row[3].value:,} 마지막 G/L 일자 : {row[2].value}')

        print('\n')

    def input_company(self) -> str:
        print('회사를 선택해주세요 \n')
        return input()

    def input_date_GL(self) -> str:
        print('G/L 일자를 입력하세요(포맷: 6자리 연월일) ex) 201010 \n')
        date_GL_selected = input()

        if len(date_GL_selected) == 6:
            date_GL_selected = f'{date_GL_selected[0:2]}/{date_GL_selected[2:4]}/{date_GL_selected[4:6]}'
            return date_GL_selected

        return self.input_date_GL()

    def get_selected_sheet(self, company_selected: str) -> Worksheet:
        sheet = self.wb['Sheet1']
        return sheet[int(company_selected)+1]


if __name__ == "__main__":
    excel = Excel()
