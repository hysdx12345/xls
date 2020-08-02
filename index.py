import xlrd
import sshtunnel
from mysqlhelper import *

filename = "异常跟踪记录表20200728.xlsx"

data = xlrd.open_workbook(filename)
sheetNames = data.sheet_names()
ProblemList = []

for item in sheetNames:
    sheet = data.sheet_by_name(item)
    sheetNrows = sheet.nrows


    class Problem:
        pass


    for i in range(sheetNrows):
        if i > 0:
            obj = Problem()  # 构建问题对象
            obj.Num = sheet.cell(i, 0).value
            obj.Date = xlrd.xldate_as_datetime(sheet.cell(i, 1).value, 0).date()
            obj.Name = sheet.cell(i, 2).value
            obj.Problem = sheet.cell(i, 5).value
            ProblemList.append(obj)
    sheet = data.sheet_by_name(item)
    sheetNrows = sheet.nrows

# mydb = dbhelper("localhost", 3306, "root", "123456", "test")
# mydb.excutedata("DROP TABLE IF EXISTS problem_tbl")
# mydb.excutedata(
#     "CREATE TABLE IF NOT EXISTS problem_tbl (id INT UNSIGNED NOT NULL AUTO_INCREMENT,problem_date DATE,product_name VARCHAR(255),problem VARCHAR(255),PRIMARY KEY (`id`))")
#
# sql = "INSERT INTO problem_tbl (problem_date, product_name, problem) VALUES (%s, %s, %s)"
# vals = []
# for item in ProblemList:
#     vals.append((item.Date, item.Name, item.Problem))
# mydb.excutemanydata(sql, vals)

with sshtunnel.SSHTunnelForwarder(
        ('122.51.129.15', 22),
        ssh_username='root',
        ssh_password='han910215@',
        remote_bind_address=('localhost', 3306),
        local_bind_address=('127.0.0.1', 13306)
) as tunnel:
    mydb = dbhelper("localhost", 13306, "root", "123456", "test")
    mydb.excutedata("DROP TABLE IF EXISTS problem_tbl")
    mydb.excutedata(
        "CREATE TABLE IF NOT EXISTS problem_tbl (id INT UNSIGNED NOT NULL AUTO_INCREMENT,problem_date DATE,product_name VARCHAR(255),problem VARCHAR(255),PRIMARY KEY (`id`))")

    sql = "INSERT INTO problem_tbl (problem_date, product_name, problem) VALUES (%s, %s, %s)"
    vals = []
    for item in ProblemList:
        vals.append((item.Date, item.Name, item.Problem))
    mydb.excutemanydata(sql, vals)