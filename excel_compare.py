import xlrd
import os
import xlsxwriter

class ExcelMeta:
    def __init__(self, excel_path, worksheet, ignore_list):
        self.excel_path = excel_path
        self.worksheet = worksheet
        self.ignore_list = ignore_list

    def __str__(self):
        return self.excel_path + ' - ' +self.worksheet

class ExcelCompare:
    def __init__(self, excel_meta_a, excel_meta_b, export_file_path, export_sheet_name='중복데이터'):
        self.excel_meta_a = excel_meta_a
        self.excel_meta_b = excel_meta_b
        self.export_file_path = export_file_path
        self.export_sheet_name = export_sheet_name

    def analyze(self):
        print('1. load excel sheet start')
        print('  ', str(self.excel_meta_a))
        print('  ', str(self.excel_meta_b))
        print()
        self.load_excel_sheet()

        print('2. find intersection id start')
        intersection_id_set = self.find_intersection_id()
        print('  intersection id count - ', len(intersection_id_set))
        print('  ', intersection_id_set)
        print()

        print('3. export excel - ', self.export_file_path)
        total_row = self.export_excel(intersection_id_set)
        print('  total row for intersaction id: ', total_row)
        print()
        print('4. finish >_<!!!!!!!')
        print('-' * 40)
        print()


    def load_excel_sheet(self):
        book = xlrd.open_workbook(self.excel_meta_a.excel_path)
        self.sheet_a = book.sheet_by_name(self.excel_meta_a.worksheet)

        book = xlrd.open_workbook(self.excel_meta_b.excel_path)
        self.sheet_b = book.sheet_by_name(self.excel_meta_b.worksheet)

    def find_intersection_id(self):
        set_a_id = set()

        for i in range(1, self.sheet_a.nrows):
            set_a_id.add(self.sheet_a.cell_value(rowx=i, colx=1))

        set_b_id = set()

        for i in range(1, self.sheet_b.nrows):
            set_b_id.add(self.sheet_b.cell_value(rowx=i, colx=1))

        return set_a_id & set_b_id

    def export_excel(self, intersection_id_set):
        if os.path.exists(self.export_file_path):
            os.remove(self.export_file_path)
        workbook = xlsxwriter.Workbook(self.export_file_path)
        worksheet = workbook.add_worksheet(self.export_sheet_name)

        for i in range(self.sheet_a.ncols):
            worksheet.write(0, i, self.sheet_a.cell_value(rowx=0, colx=i))

        current_row = 1

        for i in range(1, self.sheet_a.nrows):
            if self.sheet_a.cell_value(rowx=i, colx=1) not in intersection_id_set:
                continue
            current_col_index = 0
            for j in range(self.sheet_a.ncols):
                if j in self.excel_meta_a.ignore_list:
                    continue
                worksheet.write(current_row, current_col_index,
                                self.sheet_a.cell_value(rowx=i, colx=j))
                current_col_index += 1
            current_row += 1

        for i in range(1, self.sheet_b.nrows):
            if self.sheet_b.cell_value(rowx=i, colx=1) not in intersection_id_set:
                continue
            current_col_index = 0
            for j in range(self.sheet_b.ncols):
                if j in self.excel_meta_b.ignore_list:
                    continue
                worksheet.write(current_row, current_col_index,
                                self.sheet_b.cell_value(rowx=i, colx=j))
                current_col_index += 1
            current_row += 1


        workbook.close()
        return current_row