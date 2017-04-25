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
        print('-' * 40)
        print('1. load excel sheet start')
        print('  ', str(self.excel_meta_a))
        print('  ', str(self.excel_meta_b))
        print()
        self.load_excel_sheet()

        print('2. find intersection id start')
        intersection_id_set, id_dic_a, id_dic_b \
            = self.find_intersection_id()
        print('  intersection id count - ', len(intersection_id_set))
        print('  ', intersection_id_set)
        print()

        print('3. export excel - ', self.export_file_path)
        total_row = self.export_excel(intersection_id_set, id_dic_a, id_dic_b)
        print('  total row for intersaction id: ', total_row)
        print()
        print('4. finish >_<!!!!!!!')
        print('-' * 40)


    def load_excel_sheet(self):
        book = xlrd.open_workbook(self.excel_meta_a.excel_path)
        self.sheet_a = book.sheet_by_name(self.excel_meta_a.worksheet)

        book = xlrd.open_workbook(self.excel_meta_b.excel_path)
        self.sheet_b = book.sheet_by_name(self.excel_meta_b.worksheet)

    def find_intersection_id(self):
        id_dic_a = {}
        set_a_id = set()

        for i in range(1, self.sheet_a.nrows):
            count_learning = self.sheet_a.cell_value(rowx=i, colx=self.sheet_a.ncols - 1)
            if str(count_learning) == '0':
                continue
            id = self.sheet_a.cell_value(rowx=i, colx=1)
            set_a_id.add(id)
            if id not in id_dic_a:
                id_dic_a[id] = 0
            id_dic_a[id] += 1
        id_dic_b = {}
        set_b_id = set()

        for i in range(1, self.sheet_b.nrows):
            count_learning = self.sheet_b.cell_value(rowx=i, colx=self.sheet_b.ncols - 1)
            if str(count_learning) == '0':
                continue
            id = self.sheet_b.cell_value(rowx=i, colx=1)
            set_b_id.add(id)
            if id not in id_dic_b:
                id_dic_b[id] = 0
            id_dic_b[id] += 1

        return (set_a_id & set_b_id), id_dic_a, id_dic_b

    def export_excel(self, intersection_id_set, id_dic_a, id_dic_b):
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
            count_learning = self.sheet_a.cell_value(rowx=i, colx=self.sheet_a.ncols - 1)
            if str(count_learning) == '0':
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
            count_learning = self.sheet_b.cell_value(rowx=i, colx=self.sheet_b.ncols - 1)
            if str(count_learning) == '0':
                continue
            current_col_index = 0
            for j in range(self.sheet_b.ncols):
                if j in self.excel_meta_b.ignore_list:
                    continue
                worksheet.write(current_row, current_col_index,
                                self.sheet_b.cell_value(rowx=i, colx=j))
                current_col_index += 1
            current_row += 1

        worksheet_id_dic_a = workbook.add_worksheet(self.excel_meta_a.worksheet)

        index = 0
        for key, value in id_dic_a.items():
            worksheet_id_dic_a.write(index, 0, key)
            worksheet_id_dic_a.write(index, 1, value)
            index += 1

        worksheet_id_dic_b = workbook.add_worksheet(self.excel_meta_b.worksheet)

        index = 0
        for key, value in id_dic_b.items():
            worksheet_id_dic_b.write(index, 0, key)
            worksheet_id_dic_b.write(index, 1, value)
            index += 1

        workbook.close()
        return current_row