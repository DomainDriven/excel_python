# -*- coding: cp949 -*-
import xlwings as xw
#import pythoncom

def excel_col(num):
	arr = []
	while num > 0:
		num -= 1
		col = num % 26
		arr.append(unichr(ord('A') + col))
		num /= 26
	return "".join(arr[::-1])

def get_match_row(value, col_list, row_num):
	for row in xrange(1,row_num+1):
		for col_idx in col_list:
			if col_idx >= 1:
				xy = excel_col(col_idx) + str(row)
				if xw.Range(xy).value == value:
					return row
	return -1
		
#VB Sub �������_���()
def project_list(file):
	# �۾� ����: ���� ���Ǵ�� ǥ���� �����Ǵ� ��� ���� ���� ����� �μ��Ѵ�..
	# ��Ʈ Ÿ��: ��Ʈ�̸��� ���ڸ� ��� ��Ʈ��, ���ڸ� �������� ��Ʈ�� �����Ѵ�.
	# Test ���: shtCount ������ ���� ���� �־� �׽�Ʈ �� �� �� ��ü�� �����Ѵ�.

	print file		
	wb = xw.Workbook(file)

	# �� �� ����.
	sumShtName = "�������"	
	# �� �� ����.
	formShtName = "���뼭��"
	# �� �� ����.
	shtCount = 20                 # �׽�Ʈ �� ��Ʈ ��... =20-12\
	#VB ''shtCount = Sheets.Count   ''��� �� ��Ʈ ��... �׽�Ʈ ������ �ڸ�Ʈ ���·�
	# shtCount = xw.Sheet.count() # ��� �� ��Ʈ ��... �׽�Ʈ ������ �ڸ�Ʈ ���·�

	#VB If shtCount > Sheets.Count Then
	#VB    shtCount = Sheets.Count
	#VB End If
	if shtCount > xw.Sheet.count():
		shtCount = xw.Sheet.count()

        print "��¥ ��Ʈ ���� : ", xw.Sheet.count(), "�׽�Ʈ ��Ʈ ���� : " , shtCount

	#VB ''--------------------------------------------------------------------------------
	# --------------------------------------------------------------------------------

	#VB Dim nameRow As Integer: nameRow = 1    ''ǥ�̸�
	nameRow = 1    # ǥ�̸�
	#VB Dim yearRow As Integer: yearRow = 2    ''��⿬��
	yearRow = 2    # ��⿬��

	#VB Dim supplyRow As Integer        ''����ǥ
	#VB Dim demandRow As Integer        ''����ǥ

	#VB Dim listRow As Integer          ''���ǥ
	#VB Dim listContRow As Integer      ''��� ����

	#VB Dim i As Integer
	#VB Dim j As Integer
 
	#VB Dim usedRows As Integer 
	#VB Dim projName As String
   
	#VB Dim sumShtCount As Integer   ''��Ʈ �̸� ����, non-numeric

	#VB Dim shtName As String

	#VB Dim sumSht As Worksheet    ''��� ��Ʈ
	#VB Dim formSht As Worksheet
	#VB Dim ws As Worksheet         ''���� ��Ʈ

	#VB     ''����� �ʱ�ȭ, Run time ����
	#VB Sheets(formShtName).Activate
	#VB Set formSht = Sheets(formShtName)
	xw.Sheet(formShtName).activate()
	formSht = xw.Sheet(formShtName)
	
	#VB Sheets(sumShtName).Activate
	#VB Set sumSht = Sheets(sumShtName)
	xw.Sheet(sumShtName).activate()
	sumSht = xw.Sheet(sumShtName)

	#VB Set sumSht = Application.ActiveSheet
	# sumSht = Application.ActiveSheet <= �ּ� Ǯ��� ��. <= �� �ᵵ �� ��

	#VB usedRows = sumSht.UsedRange.Rows.Count
	# https://github.com/ZoomerAnalytics/xlwings/issues/112
	usedRows = len(sumSht.xl_sheet.UsedRange.Rows)

	#VB supplyRow = WorksheetFunction.Match("����ǥ", sumSht.Range(Cells(1, 2), Cells(usedRows, 2)), 0)
	#VB demandRow = WorksheetFunction.Match("����ǥ", sumSht.Range(Cells(1, 2), Cells(usedRows, 2)), 0)
	#VB listRow = WorksheetFunction.Match("���ǥ", sumSht.Range(Cells(1, 2), Cells(usedRows, 2)), 0)
	#VB listContRow = WorksheetFunction.Match("��Ͻ�1", sumSht.Range(Cells(1, 2), Cells(usedRows, 2)), 0)
	supplyRow = get_match_row(u"����ǥ", [1,2], usedRows)
	demandRow = get_match_row(u"����ǥ", [1,2], usedRows)
	listRow = get_match_row(u"���ǥ", [1,2], usedRows)
	listContRow = get_match_row(u"��Ͻ�1", [1,2], usedRows)
	#print supplyRow
	#print demandRow
	#print listRow
	#print listContRow

	#VB ''���ǥ �Ӹ��ٰ� ���� ���� ����
	#VB sumSht.Range("A" & listRow + 1 & ":AC" & listContRow - 1).Clear
	xw.Range("A"+str(listRow+1)+":AC"+str(listContRow - 1)).clear()

	# �� �� ����.
	sumShtCount = 0
	
	#VB For i = 1 To shtCount
	for i in xrange(1, shtCount+1):
		#VB shtName = Sheets.Item(i).Name
		shtName = xw.Sheet(i).name
		print shtName
		
		#VB ''�ٸ� ������ ��Ʈ���� �ٿ��ֱ� ���ϰ�,
		#VB ''���� ��Ʈ���� �ٿ��ֱ� �Ѵ�
		#VB If Not IsNumeric(shtName) Then
		#VB     sumShtCount = sumShtCount + 1
		#VB     GoTo Next_wsBuf
                #VB End If
		if not shtName.isdigit():
			sumShtCount += 1
			continue

		#VB Set ws = Sheets(shtName)
		#ws = xw.Sheet(shtName)

		#VB ws.Range("A" & nameRow & ":AC" & nameRow + 8).Clear
		#VB ''������ ����� I/F Buf�� �����,
		#VB ''�� 9�� �����.
		xw.Range(shtName, "A"+str(nameRow)+":AC"+str(nameRow+8)).clear()

        	#VB ''������Ʈ ���� ���ǥ �̸��� �����.
		#VB sumSht.Range("B2") = sumShtName  ''B2: ���ǥ �̸�
		xw.Range(sumShtName, "B2").value = sumShtName
		print xw.Range(sumShtName, "B2").value, sumShtName

		#VB sumSht.Range("A" & nameRow & ":AC" & yearRow).Copy _
		#VB     ws.Range("A" & nameRow)
		xw.Range(shtName,"A"+str(nameRow)).value = xw.Range(sumShtName,"A"+str(nameRow)+":AC"+str(yearRow)).value

		#VB sumSht.Range("A" & supplyRow & ":AC" & supplyRow + 2).Copy _
		#VB ws.Range("A" & supplyRow)
		# sdr1982 - ����ǥ�� �Ʒ� �κ�(supplyRow+1)�� ���� ������ �� �κ�(supplyRow)�� ������ ����.
		xw.Sheet(shtName).activate()
		xw.Range(shtName,"A"+str(supplyRow+1)).value = xw.Range(sumShtName,"A"+str(supplyRow+1)+":AC"+str(supplyRow+2)).value
		xw.Range(shtName,"A"+str(supplyRow)+":AC"+str(supplyRow+1)).formula = xw.Range(sumShtName,"A"+str(supplyRow)+":AC"+str(supplyRow+1)).formula

		#VB ''���뼭��.�����Ⱓ,���Ⱓ
		#VB formSht.Range("B10:R15").Copy ws.Range("B10")
		xw.Range(shtName,"B10").value = xw.Range(formShtName,"B10:R15").value

        	#VB ''Form��Ʈ���� �Լ��� �� ����, ���İ� ���� C&P
        	#VB ''formSht.Range("C1:T2").Copy   ''�� �� copy�ؼ� ���� �ٸ� ����� ��󾴴�.
        	#VB ''ws.Range("C1").PasteSpecial xlPasteFormats
        	#VB ''ws.Range("C1").PasteSpecial xlPasteFormulas
              
	#VB Next_wsBuf:
	#VB Next i

	#VB '' ������ȣ ��Ʈ ����� ���� ��Ʈ C���� �ۼ��Ѵ�.
	#VB '' �ٸ� ������ ��Ʈ���� �״�� ����Ѵ�.
	#VB '' �ۼ����� ������ ��Ʈ���� ���� ��� ������ ���� ��, �ٽ� ä���
    
	#VB sumSht.Range("C" & listContRow & ":AC" & usedRows).Clear
	xw.Sheet(sumShtName).activate()
	xw.Range(sumShtName, "C"+str(listContRow)+":AC"+str(usedRows)).clear()

	# �� �� ����.
	rowCount = listContRow      #''

	#VB For i = 1 To shtCount
	shtSumCount = 0
	for i in xrange(1, shtCount+1):
		#VB shtName = Sheets.Item(i).Name
		shtName = xw.Sheet(i).name
     
		#VB If Not IsNumeric(shtName) Then
		#VB	sumShtCount = sumShtCount + 1
		#VB GoTo Next_wsBuf
		#VB End If
		if not shtName.isdigit():
			shtSumCount += 1
			continue
     
		#VB ''������ȣ ���� ���İ� �� �ٿ��ֱ�
		#VB ''������ȣ ���� ���� ���� �ٿ��ֱ�
		#VB sumSht.Range("C" & demandRow + 1).Copy     ''�ߴ���. ��Ͻ�1> ������ȣ
		#VB sumSht.Range("C" & rowCount + 0).PasteSpecial xlPasteFormats
		xw.Range(sumShtName,"C"+str(rowCount)).value = xw.Range(sumShtName, "C"+str(demandRow+1)).value
        
		#VB ''sumSht.Range("C" & demandRow + 2).Copy     ''��Ͻ�2> ������ȣ
		#VB ''sumSht.Range("C" & rowCount + 1).PasteSpecial xlPasteFormats
        
		#VB ''�׸��� ������ȣ ���� �� �ٿ��ֱ�
		#VB sumSht.Range("C" & rowCount + 0) = shtName
		xw.Range(sumShtName,"C"+str(rowCount)).value = shtName
		#VB ''sumSht.Range("C" & rowCount + 1) = shtName
		#VB ''sumSht.Range("C" & rowCount + 2) = ""
		
		# �� �� ����.
		rowCount = rowCount + 1
     
	#VB NextPIDList:
	#VB Next i

	#VB ''������ ���� ���� �ٿ��ֱ�
    
	#VB sumSht.Range("B" & listRow) = "���ǥ"
	xw.Range(sumShtName,"B"+str(listRow)).value = "���ǥ"
	print "B"+str(listRow), "���ǥ"
	#VB sumSht.Range("A" & listContRow) = "listContRow"
	xw.Range(sumShtName, "A"+str(listContRow)).value = "listContRow"
	print "A"+str(listContRow), "listContRow"
    
	#VB sumSht.Range("B" & listRow).Font.Color = RGB(255, 0, 0)
	#xw.Range(sumShtName,"B"+str(listRow)).color = (255,0,0)
	#VB sumSht.Range("B" & listRow).Font.Size = 12
	#VB sumSht.Range("B" & listRow).HorizontalAlignment = xlCenter
    
	#VB '' ������ �Ӹ����� �� ���� �μ��Ѵ�.
	#VB sumSht.Range("C" & demandRow & ":U" & demandRow).Copy
	#VB sumSht.Range("C" & listRow).PasteSpecial (xlPasteAll)
	xw.Range(sumShtName,"C"+str(listRow)).value = xw.Range(sumShtName, "C"+str(demandRow)+":U"+str(demandRow)).value
    
	#VB ''--------------------------------------------------------------------------------

	#VB ''����ǥ�� �����1�� ���ǥ�� ��Ͻ�1�� �ٿ��ֱ� �Ѵ�.
      
	#VB sumSht.Range("D" & demandRow + 1 & ":U" & demandRow + 1).Copy
	#VB sumSht.Range("D" & listContRow + 0).PasteSpecial xlPasteAllUsingSourceTheme
	# TODO - ���� ���� �Լ�
	xw.Range(sumShtName,"D"+str(listContRow)+":U"+str(listContRow)).formula = xw.Range(sumShtName, "D"+str(demandRow+1)+":U"+str(demandRow+1)).formula
	#print sumShtName, "D"+str(demandRow+1)+":U"+str(demandRow+1), xw.Range(sumShtName,"D"+str(demandRow+1)+":U"+str(demandRow+1)).formula


	#src_tuples = xw.Range(sumShtName,"D"+str(demandRow+1)+":U"+str(demandRow+1)).formula
	#tar = []
	#for tuple in src_tuples:
	#	for s in tuple:
	#		tar.append(s.replace("C11","C"+str(listContRow)))	
	#xw.Range(sumShtName,"D"+str(listContRow)+":U"+str(listContRow)).formula = tar
         
	#VB ''sumSht.Range("D" & demandRow + 2 & ":U" & demandRow + 2).Copy
	#VB ''sumSht.Range("D" & listContRow + 1).PasteSpecial xlPasteAllUsingSourceTheme        
	#VB ''sumSht.Range("D" & listContRow + 2) = ""

	# �� �� ����.
	rowCount = listContRow + 1
	#VB For i = shtSumCount + 1 To shtCount  ''PID ù���� �̹� ���, n-1�� �߰�
	print "shtSumCount", shtSumCount, "shtCount", shtCount
	for i in xrange(shtSumCount + 1, shtCount+1):
		#VB shtName = Application.Sheets(i).Name
		shtName = xw.Sheet(i).name

		#VB If Not IsNumeric(shtName) Then
		#VB 	GoTo NextListCont
		#VB End If
		if not shtName.isdigit():
			continue

		# ���� �� ���� �������� �����ϴ� ��� �� �ʿ� ���� ����

		#VB ''ù�� �ν��Ͻ����� ���İ� ������ �����ؼ� ��Ʈ ���� ��ŭ �ٿ��ֱ��Ѵ�            
		#VB ''�� ���İ� ���� �ٿ��ֱ� (������ȣ�� �ռ� �Ϸ�)
		#VB sumSht.Range("D" & listContRow & ":U" & listContRow).Copy
        	#VB sumSht.Range("D" & rowCount).PasteSpecial xlPasteFormats
		#try:
		#	xw.Application(wb).xl_app.Run('sumShtName = "�������"')
		#	xw.Application(wb).xl_app.Run('Set sumSht = Sheets(sumShtName)')
		#	xw.Application(wb).xl_app.Run('sumSht.Range("D" & listContRow & ":U" & listContRow).Copy')
		#	xw.Application(wb).xl_app.Run('sumSht.Range("D" & rowCount).PasteSpecial xlPasteFormats')
		#except pythoncom.com_error as e:
		#	print e
		#	print e[0]
		#	print e[1]
		#	print e[2]
		#	print e[2][2]
		#	break	
		#xw.Range(sumShtName,"D"+str(rowCount)).value = xw.Range(sumShtName, "D"+str(listContRow)+":U"+str(listContRow)).value
        
		#VB ''�� ���� �ٿ��ֱ� (������ȣ�� �ռ� �Ϸ�)
		#VB ''sumSht.Range("D" & listContRow + 0 & ":U" & listContRow + 0).Copy
		#VB sumSht.Range("D" & rowCount).PasteSpecial xlPasteFormulas
		#print sumShtName, "D"+str(demandRow+1)+":U"+str(demandRow+1), xw.Range(sumShtName,"D"+str(listContRow)+":U"+str(listContRow)).formula

		src_tuples = xw.Range(sumShtName,"D"+str(listContRow)+":U"+str(listContRow)).formula
		tar = []
		for tuple in src_tuples:
			for s in tuple:
				tar.append(s.replace("C11","C"+str(rowCount)))	
		xw.Range(sumShtName,"D"+str(rowCount)+":U"+str(rowCount)).formula = tar

		#print sumShtName, "D"+str(rowCount), xw.Range(sumShtName,"D"+str(rowCount)).formula
		#print sumShtName, "D"+str(rowCount), xw.Range(sumShtName,"D"+str(rowCount)).value
		#print sumShtName, "D"+str(listContRow), xw.Range(sumShtName,"D"+str(listContRow)+":U"+str(listContRow)).formula
        
		# �� �� ����.
        	rowCount = rowCount + 1
     
	#VB NextListCont:
	#VB Next i

# TODO - ���� ���� �Լ� 
#def copy_fomula(src_name, src_row, tar_name, tar_row):
#	src_tuples = xw.Range(src_name,"D"+str(src_row)+":U"+str(src_row)).formula

if __name__ == "__main__":
	import os 
	cur_dir = os.path.dirname(os.path.realpath(__file__))
	project_list(cur_dir + os.path.sep +'test.xls')










