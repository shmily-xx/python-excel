import xlrd
import xlwt
from xlrd import xldate_as_tuple

import xlsxwriter

import sys

# 导入需要读取的第一个Excel表格的路径

data = xlrd.open_workbook(r'D:\\code\\python-excel-handle\\etcOra.xlsx')

table = data.sheets()[1]

# 创建一个空列表，存储Excel的数据

tables = []

# 将excel表格内容导入到tables列表中


def import_excel(excel):

 #   for rown in range(5):
 for rown in range(excel.nrows):

  #  array = {'costTitle': '', 'entryTime': '', 'entryPort': '', 'leaveTime': '', 'leavePort': ''}
   array = {'entryTime': '', 'entryPort': '', 'leaveTime': '', 'leavePort': ''}

   str = table.cell_value(rown, 0)
  #  array['costTitle'] = str
   titleArr = str.split('|')
   array['entryTime'] = titleArr[0]
   array['entryPort'] = titleArr[1]
   array['leaveTime'] = titleArr[2]
   array['leavePort'] = titleArr[3]

   tables.append(array)


def write_excel(tables):
    print("-"*50)

    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('ssss', cell_overwrite_ok=True)
    # 循环写入表头
    title = ["进站时间", "进站地点", "出站时间", "出站地点"]
    for t in title:
        # 写入excel
        worksheet.write(0, 0, t)

    i = 1
    n = 0
    while n <= len(tables) - 1:
        lst = list(tables[n].values())
        for t in lst:
          worksheet.write(i, n, t)
        print(n/len(tables))
        n += 1
    i += 1

    workbook.save(r'D:\\code\\python-excel-handle\\etc.xlsx')


def write(tables):
  workbook = xlsxwriter.Workbook('new_excel.xlsx')  # 新建excel表


  worksheet = workbook.add_worksheet('sheet1')  # 新建sheet（sheet的名称为"sheet1"）

  headings = ["进站时间", "进站地点", "出站时间", "出站地点"]  # 设置表头

  worksheet.write_row(0, 0, headings)

  i = 1
  n = 0
  while n <= len(tables) - 1:
    progress_bar(n, len(tables))
    lst = list(tables[n].values())
    worksheet.write_row(i, 0, lst)
    n += 1
    i += 1
  

  workbook.close()  # 将excel文件保存关闭，如果没有这一行运行代码会报错

def progress_bar(n, l):

  print(n, l, sep="/")

  sys.stdout.flush()


if __name__ == '__main__':

  # 将excel表格的内容导入到列表中

  import_excel(table)

  # write_excel(tables)
  write(tables)

  # for o in tables:
  #   print(o)
