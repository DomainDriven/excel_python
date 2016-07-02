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
		
#VB Sub 과제목록_당기()
def project_list(file):
	# 작업 내용: 현재 예실대비 표에서 유지되는 당기 진행 과제 목록을 인쇄한다..
	# 시트 타입: 시트이름이 문자면 요약 시트고, 숫자면 개별과제 시트로 구별한다.
	# Test 방법: shtCount 변수에 적은 값을 넣어 테스트 해 본 후 전체를 실행한다.

	print file		
	wb = xw.Workbook(file)

	# 둘 다 같음.
	sumShtName = "과제목록"	
	# 둘 다 같음.
	formShtName = "공통서식"
	# 둘 다 같음.
	shtCount = 20                 # 테스트 할 시트 수... =20-12\
	#VB ''shtCount = Sheets.Count   ''당기 총 시트 수... 테스트 동안은 코멘트 상태로
	# shtCount = xw.Sheet.count() # 당기 총 시트 수... 테스트 동안은 코멘트 상태로

	#VB If shtCount > Sheets.Count Then
	#VB    shtCount = Sheets.Count
	#VB End If
	if shtCount > xw.Sheet.count():
		shtCount = xw.Sheet.count()

        print "진짜 시트 개수 : ", xw.Sheet.count(), "테스트 시트 개수 : " , shtCount

	#VB ''--------------------------------------------------------------------------------
	# --------------------------------------------------------------------------------

	#VB Dim nameRow As Integer: nameRow = 1    ''표이름
	nameRow = 1    # 표이름
	#VB Dim yearRow As Integer: yearRow = 2    ''당기연도
	yearRow = 2    # 당기연도

	#VB Dim supplyRow As Integer        ''공급표
	#VB Dim demandRow As Integer        ''수요표

	#VB Dim listRow As Integer          ''목록표
	#VB Dim listContRow As Integer      ''목록 내용

	#VB Dim i As Integer
	#VB Dim j As Integer
 
	#VB Dim usedRows As Integer 
	#VB Dim projName As String
   
	#VB Dim sumShtCount As Integer   ''시트 이름 갯수, non-numeric

	#VB Dim shtName As String

	#VB Dim sumSht As Worksheet    ''요약 시트
	#VB Dim formSht As Worksheet
	#VB Dim ws As Worksheet         ''개별 시트

	#VB     ''상수값 초기화, Run time 에서
	#VB Sheets(formShtName).Activate
	#VB Set formSht = Sheets(formShtName)
	xw.Sheet(formShtName).activate()
	formSht = xw.Sheet(formShtName)
	
	#VB Sheets(sumShtName).Activate
	#VB Set sumSht = Sheets(sumShtName)
	xw.Sheet(sumShtName).activate()
	sumSht = xw.Sheet(sumShtName)

	#VB Set sumSht = Application.ActiveSheet
	# sumSht = Application.ActiveSheet <= 주석 풀어야 함. <= 안 써도 될 듯

	#VB usedRows = sumSht.UsedRange.Rows.Count
	# https://github.com/ZoomerAnalytics/xlwings/issues/112
	usedRows = len(sumSht.xl_sheet.UsedRange.Rows)

	#VB supplyRow = WorksheetFunction.Match("공급표", sumSht.Range(Cells(1, 2), Cells(usedRows, 2)), 0)
	#VB demandRow = WorksheetFunction.Match("수요표", sumSht.Range(Cells(1, 2), Cells(usedRows, 2)), 0)
	#VB listRow = WorksheetFunction.Match("목록표", sumSht.Range(Cells(1, 2), Cells(usedRows, 2)), 0)
	#VB listContRow = WorksheetFunction.Match("목록식1", sumSht.Range(Cells(1, 2), Cells(usedRows, 2)), 0)
	supplyRow = get_match_row(u"공급표", [1,2], usedRows)
	demandRow = get_match_row(u"수요표", [1,2], usedRows)
	listRow = get_match_row(u"목록표", [1,2], usedRows)
	listContRow = get_match_row(u"목록식1", [1,2], usedRows)
	#print supplyRow
	#print demandRow
	#print listRow
	#print listContRow

	#VB ''목록표 머리줄과 내용 사이 지움
	#VB sumSht.Range("A" & listRow + 1 & ":AC" & listContRow - 1).Clear
	xw.Range("A"+str(listRow+1)+":AC"+str(listContRow - 1)).clear()

	# 둘 다 같음.
	sumShtCount = 0
	
	#VB For i = 1 To shtCount
	for i in xrange(1, shtCount+1):
		#VB shtName = Sheets.Item(i).Name
		shtName = xw.Sheet(i).name
		print shtName
		
		#VB ''다른 공통요약 시트에는 붙여넣기 안하고,
		#VB ''개별 시트에만 붙여넣기 한다
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
		#VB ''이전에 사용한 I/F Buf를 지운다,
		#VB ''총 9행 지운다.
		xw.Range(shtName, "A"+str(nameRow)+":AC"+str(nameRow+8)).clear()

        	#VB ''개별시트 마다 요약표 이름을 등록함.
		#VB sumSht.Range("B2") = sumShtName  ''B2: 요약표 이름
		xw.Range(sumShtName, "B2").value = sumShtName
		print xw.Range(sumShtName, "B2").value, sumShtName

		#VB sumSht.Range("A" & nameRow & ":AC" & yearRow).Copy _
		#VB     ws.Range("A" & nameRow)
		xw.Range(shtName,"A"+str(nameRow)).value = xw.Range(sumShtName,"A"+str(nameRow)+":AC"+str(yearRow)).value

		#VB sumSht.Range("A" & supplyRow & ":AC" & supplyRow + 2).Copy _
		#VB ws.Range("A" & supplyRow)
		# sdr1982 - 공급표의 아래 부분(supplyRow+1)은 값을 보내고 윗 부분(supplyRow)은 수식을 보냄.
		xw.Sheet(shtName).activate()
		xw.Range(shtName,"A"+str(supplyRow+1)).value = xw.Range(sumShtName,"A"+str(supplyRow+1)+":AC"+str(supplyRow+2)).value
		xw.Range(shtName,"A"+str(supplyRow)+":AC"+str(supplyRow+1)).formula = xw.Range(sumShtName,"A"+str(supplyRow)+":AC"+str(supplyRow+1)).formula

		#VB ''공통서식.과제기간,당기기간
		#VB formSht.Range("B10:R15").Copy ws.Range("B10")
		xw.Range(shtName,"B10").value = xw.Range(formShtName,"B10:R15").value

        	#VB ''Form시트에서 함수형 셀 복사, 서식과 수식 C&P
        	#VB ''formSht.Range("C1:T2").Copy   ''한 번 copy해서 각기 다른 기능을 골라쓴다.
        	#VB ''ws.Range("C1").PasteSpecial xlPasteFormats
        	#VB ''ws.Range("C1").PasteSpecial xlPasteFormulas
              
	#VB Next_wsBuf:
	#VB Next i

	#VB '' 과제번호 시트 목록을 개별 시트 C열에 작성한다.
	#VB '' 다른 공통요약 시트들은 그대로 통과한다.
	#VB '' 작성중인 공통요약 시트에서 이전 목록 내용을 지운 후, 다시 채운다
    
	#VB sumSht.Range("C" & listContRow & ":AC" & usedRows).Clear
	xw.Sheet(sumShtName).activate()
	xw.Range(sumShtName, "C"+str(listContRow)+":AC"+str(usedRows)).clear()

	# 둘 다 같음.
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
     
		#VB ''과제번호 열에 서식과 값 붙여넣기
		#VB ''과제번호 열에 서식 먼저 붙여넣기
		#VB sumSht.Range("C" & demandRow + 1).Copy     ''중단점. 목록식1> 과제번호
		#VB sumSht.Range("C" & rowCount + 0).PasteSpecial xlPasteFormats
		xw.Range(sumShtName,"C"+str(rowCount)).value = xw.Range(sumShtName, "C"+str(demandRow+1)).value
        
		#VB ''sumSht.Range("C" & demandRow + 2).Copy     ''목록식2> 과제번호
		#VB ''sumSht.Range("C" & rowCount + 1).PasteSpecial xlPasteFormats
        
		#VB ''그리고 과제번호 열에 값 붙여넣기
		#VB sumSht.Range("C" & rowCount + 0) = shtName
		xw.Range(sumShtName,"C"+str(rowCount)).value = shtName
		#VB ''sumSht.Range("C" & rowCount + 1) = shtName
		#VB ''sumSht.Range("C" & rowCount + 2) = ""
		
		# 둘 다 같음.
		rowCount = rowCount + 1
     
	#VB NextPIDList:
	#VB Next i

	#VB ''공통요약 셀에 서식 붙여넣기
    
	#VB sumSht.Range("B" & listRow) = "목록표"
	xw.Range(sumShtName,"B"+str(listRow)).value = "목록표"
	print "B"+str(listRow), "목록표"
	#VB sumSht.Range("A" & listContRow) = "listContRow"
	xw.Range(sumShtName, "A"+str(listContRow)).value = "listContRow"
	print "A"+str(listContRow), "listContRow"
    
	#VB sumSht.Range("B" & listRow).Font.Color = RGB(255, 0, 0)
	#xw.Range(sumShtName,"B"+str(listRow)).color = (255,0,0)
	#VB sumSht.Range("B" & listRow).Font.Size = 12
	#VB sumSht.Range("B" & listRow).HorizontalAlignment = xlCenter
    
	#VB '' 공통요약 머리줄은 한 번만 인쇄한다.
	#VB sumSht.Range("C" & demandRow & ":U" & demandRow).Copy
	#VB sumSht.Range("C" & listRow).PasteSpecial (xlPasteAll)
	xw.Range(sumShtName,"C"+str(listRow)).value = xw.Range(sumShtName, "C"+str(demandRow)+":U"+str(demandRow)).value
    
	#VB ''--------------------------------------------------------------------------------

	#VB ''수요표의 수요식1를 목록표의 목록식1에 붙여넣기 한다.
      
	#VB sumSht.Range("D" & demandRow + 1 & ":U" & demandRow + 1).Copy
	#VB sumSht.Range("D" & listContRow + 0).PasteSpecial xlPasteAllUsingSourceTheme
	# TODO - 수식 복사 함수
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

	# 둘 다 같음.
	rowCount = listContRow + 1
	#VB For i = shtSumCount + 1 To shtCount  ''PID 첫번은 이미 기록, n-1개 추가
	print "shtSumCount", shtSumCount, "shtCount", shtCount
	for i in xrange(shtSumCount + 1, shtCount+1):
		#VB shtName = Application.Sheets(i).Name
		shtName = xw.Sheet(i).name

		#VB If Not IsNumeric(shtName) Then
		#VB 	GoTo NextListCont
		#VB End If
		if not shtName.isdigit():
			continue

		# 다음 줄 부터 수식으로 복사하는 방법 알 필요 있음 ㄷㄷ

		#VB ''첫번 인스턴스에서 서식과 수식을 복사해서 시트 갯수 만큼 붙여넣기한다            
		#VB ''행 서식과 수식 붙여넣기 (과제번호는 앞서 완료)
		#VB sumSht.Range("D" & listContRow & ":U" & listContRow).Copy
        	#VB sumSht.Range("D" & rowCount).PasteSpecial xlPasteFormats
		#try:
		#	xw.Application(wb).xl_app.Run('sumShtName = "과제목록"')
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
        
		#VB ''행 수식 붙여넣기 (과제번호는 앞서 완료)
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
        
		# 둘 다 같음.
        	rowCount = rowCount + 1
     
	#VB NextListCont:
	#VB Next i

# TODO - 수식 복사 함수 
#def copy_fomula(src_name, src_row, tar_name, tar_row):
#	src_tuples = xw.Range(src_name,"D"+str(src_row)+":U"+str(src_row)).formula

if __name__ == "__main__":
	import os 
	cur_dir = os.path.dirname(os.path.realpath(__file__))
	project_list(cur_dir + os.path.sep +'test.xls')










