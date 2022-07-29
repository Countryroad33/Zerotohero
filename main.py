import openpyxl

def mergeform(filearr, targetfile, month): # filearr 循环数组
    target = openpyxl.load_workbook(targetfile)
    targetsheet = target.active
    for file in filearr:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        for line in ws:
            col1 = line[0].value
            col2 = line[1].value
            col3 = line[2].value
            col4 = line[3].value
            col5 = line[4].value
            col6 = line[5].value
            print(col1, col2, col3, col4, col5, col6)
            if col1 == month:
                data_list = [col1, col2, col3, col4, col5, col6] #改成变量
                targetsheet.append(data_list)
    target.save(targetfile)


if __name__ == '__main__':
    mergeform(['test1.xlsx', 'test1.xlsx'], 'test2.xlsx', 2412)

