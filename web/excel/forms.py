# -*- coding: utf-8 -*-
from django import forms

from zipfile import ZipFile
from os.path import basename
from io import BytesIO

import xlrd
import xlwt


class UploadFileForm(forms.Form):
    col = forms.CharField(max_length=50)
    excel_file = forms.FileField()

    def handleGroupExcel(self):
        file_path = '/Users/shine/Projects/python/split-excel/output'
        groupIndex = int(self.data.get('col')) - 1
        fileContent = self.files.get('excel_file')
        mem_files = {}

        # zip_file_paths = []

        wb = xlrd.open_workbook(file_contents = fileContent.read())

        sheet = wb.sheets()[0]

        rows_num = sheet.nrows
        cols_num = sheet.ncols

        filesDict = {}

        for i in range(1, rows_num):

            headers = sheet.row_values(0)
            contents = sheet.row_values(i)
            newWb = None
            sh1 = None
            contentCursor = 1

            g_content = contents[groupIndex]
            g_content = unicode(g_content) if type (g_content) is not unicode else g_content
            g_content = g_content.strip()
            if g_content in filesDict:
                contentCursor = filesDict[g_content][0]
                newWb = filesDict[g_content][1]
                sh1 = filesDict[g_content][2]
            else:
                newWb = xlwt.Workbook()
                sh1 = newWb.add_sheet('Sheet1')
                # header
                for j in range(cols_num):
                    sh1.write(0, j, headers[j])
                filesDict[g_content] = [contentCursor, newWb, sh1]

            # content
            # 序号
            sh1.write(contentCursor, 0, contentCursor)
            # 序号后的列
            for j in range(1, cols_num):
                sh1.write(contentCursor, j, contents[j])

            contentCursor += 1
            filesDict[g_content][0] = contentCursor

        for key, [contentCursor, newWb, sh] in filesDict.items():
            # excel_path = file_path + '/' + key + '.xls'
            # newWb.save(excel_path)
            f = BytesIO()
            newWb.save(f)
            f.seek(0)
            mem_files[key] = f
            # zip_file_paths.append(excel_path)

        # return self.zip_excel_files(zip_file_paths, file_path)
        return self.zip_excel_files(mem_files)


    #def zip_excel_files (self, excels, dist_path):
    def zip_excel_files (self, mem_files):

        # writing files to a zipfile
        zip_mem_file = BytesIO()
        with ZipFile(zip_mem_file, 'w') as zip:
            # writing each file one by one
            for key, mem_file in mem_files.items():
                print(key)
                zip.writestr(key + '.xls', mem_file.read())
                # in_file.close()
                # os.remove(file)

        zip_mem_file.seek(0)
        return zip_mem_file
