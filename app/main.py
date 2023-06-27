from kivy.app import App
from kivy.properties import StringProperty, BooleanProperty
from kivy.uix.widget import Widget
from kivy.uix.label import Label
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.screenmanager import ScreenManager, Screen
from os.path import join, isdir
from openpyxl import Workbook
from statistics import mean
import os


class MainUi(BoxLayout):

    main_message = StringProperty("Drop your folder here")
    root_dir = "./"
    keys = []
    sample_name = []
    ave_data_retantion_time = {}
    data_area = []
    data_amount = []
    sample_num = 0
    data_total = []
    rp_found = BooleanProperty(False)

    def __init__(self, **kwargs):
        super(MainUi, self).__init__(**kwargs)
        Window.bind(on_drop_file=self._on_file_drop)

    def _on_file_drop(self, window, file_path, x, y):

        self.main_message = file_path.decode("UTF-8")
        self.root_dir = file_path.decode("UTF-8")
        files_path_list = []

        for folder, subfolders, files in os.walk(self.root_dir):
            for f in files:
                if f.endswith("Report.TXT"):
                    files_path_list.append(os.path.join(folder, f))

        if len(files_path_list) == 0:
            self.main_message = f"NO REPORT HAS BEEN FOUND IN \n{self.root_dir}"
        else:
            self.sample_num = 0
            self.data_area = []
            self.data_amount = []
            data_retantion_time = []
            self.sample_name = []
            self.data_total = []

            for file in files_path_list:
                f = open(file, encoding="utf-16-le").readlines()
                data = [line.strip() for line in f]

                if "External Standard Report" in data:
                    self.sample_num += 1
                    starting_line = [data.index(i) for i in data if i.startswith('RetTime')][0] + 3
                    ending_line = [data.index(i) for i in data if i.startswith('Totals :')][0]

                    s_name = [line[13:] for line in data if "Sample Name: " in line][0]
                    total = [i for i in data if i.startswith('Totals :')][0].split()[2]
                    retantion_time = {line.split()[-1:][0]:float(line.split()[0:1][0]) for line in data[starting_line : ending_line]}
                    keys = []
                    values_area = []
                    values_amount = []
                    for line in data[starting_line : ending_line]:
                        keys.append(line.split()[-1:][0])
                        values_area.append([0 if line.split()[-4:-1].pop(0) == "-" else float(line.split()[-4:-1].pop(0))][0])
                        values_amount.append([0 if line.split()[-4:-1].pop(2) == "-" else float(line.split()[-4:-1].pop(2))][0])
                    d_area = dict(zip(keys, values_area))
                    d_amount = dict(zip(keys, values_amount))
                    self.sample_name.append(s_name)
                    self.data_area.append(d_area)
                    self.data_amount.append(d_amount)
                    self.data_total.append(float(total))
                    data_retantion_time.append(retantion_time)
            self.keys = keys
            print(self.keys)
            self.ave_data_retantion_time = {}
            for key in data_retantion_time[0].keys():
                self.ave_data_retantion_time[key] = mean([d[key] for d in data_retantion_time ])
            if self.sample_num != 0:
                self.rp_found = True
            self.main_message = f"{len(files_path_list)} folders contain report files of which {self.sample_num} can be converted to excel."

    

    def on_create_click(self):
        wb = Workbook()
        sheet = wb.active
        sheet.title = "Report"
        headers = self.keys
        sheet.cell(row=1, column=1).value = "Area"
        sheet.merge_cells(start_row=1, end_row=1, start_column=1, end_column=len(headers) + 1)
        sheet.cell(row=2, column=1).value = "Standards"
        for i, sample in enumerate(self.sample_name):
            sheet.cell(row=2+i+2, column=1).value = sample

        for i, header in enumerate(headers):
            sheet.cell(row=2, column= i+2).value = header
        sheet.cell(row=3, column=1).value = "Retention Time"
        column_num = 2
        for k in self.ave_data_retantion_time:
            sheet.cell(row=3, column=column_num).value = self.ave_data_retantion_time[k]
            column_num += 1

        row_num = 4
        for d in self.data_area:
            column_num = 2
            for k in d:
                sheet.cell(row=row_num, column=column_num).value = d[k]
                column_num += 1
            row_num += 1

        starting_row = self.sample_num+5
        sheet.cell(row=starting_row, column=1).value = "Amount"
        sheet.merge_cells(start_row=starting_row, end_row=starting_row, start_column=1, end_column=len(headers) + 1)
        sheet.cell(row=starting_row+1, column=1).value = "Standards"
        for i, sample in enumerate(self.sample_name):
            sheet.cell(row=starting_row+i+3, column=1).value = sample

        for i, header in enumerate(headers):
            sheet.cell(row=starting_row+1, column= i+2).value = header

        sheet.cell(row=starting_row+1, column=len(headers)+2).value = "Total"
        for i, t in enumerate(self.data_total):
            sheet.cell(row=starting_row+3+i, column=len(headers)+2).value = t

        sheet.cell(row=starting_row+2, column=1).value = "Retention Time"
        column_num = 2
        for k in self.ave_data_retantion_time:
            sheet.cell(row=starting_row+2, column=column_num).value = self.ave_data_retantion_time[k]
            column_num += 1

        row_num = starting_row + 3
        for d in self.data_amount:
            column_num = 2
            for k in d:
                sheet.cell(row=row_num, column=column_num).value = d[k]
                column_num += 1
            row_num += 1

        wb.save(f"{self.root_dir}/final_report.xlsx")
        self.main_message = "Excle file successfuly created in the same directory"




class XlsxApp(App):
    pass

if __name__ == '__main__':
    XlsxApp().run()