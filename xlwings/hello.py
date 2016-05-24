import xlwings as xw
wb = xw.Workbook()  # Creates a connection with a new workbook
xw.Range('A1').value = 'Hello World!'
xw.Range('A2').value = 'Foo 1'
#print xw.Range('A2').value
#'Foo 1'
xw.Range('A2').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
#print xw.Range('A1').table.value  # or: Range('A1:C2').value
#[['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
#xw.Sheet(1).name
#'Sheet1'
