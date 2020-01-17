import openpyxl


class Excel:
    def __init__(self):
        self.print_status()
        company_selected = self.input_company()
        date_GL_selected = self.input_date_GL()

        print(company_selected, date_GL_selected, type(date_GL_selected))

    def init_data(self) -> list:
        wb = openpyxl.load_workbook('authErp.xlsx')  # productOrder.xlsx 파일을 열어서 wb 변수에 할당
        sheet1 = wb['Sheet1']  # 엑셀의 Sheet1을 open
        return sheet1['A2':'D8']  # sheet1의 A3부터 L310까지 rows 변수에 할당

    def print_status(self):
        rows = self.init_data()
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


if __name__ == "__main__":
    excel = Excel()
