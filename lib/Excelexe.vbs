Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False

Set oWS = WScript.CreateObject("WScript.Shell")
' Get the %userprofile% in a variable, or else it won't be recognized
userProfile = oWS.ExpandEnvironmentStrings( "%userprofile%" )
xclbook = WScript.Arguments.Item(0)
its = "\Downloads\"& xclbook &".xlsx"
'msgbox(its)
'Set objWorkbook = objExcel.Workbooks.Open (userProfile & "\Downloads\Timesheet_Base_Report_SV_New.xlsx")
Set objWorkbook = objExcel.Workbooks.Open (userProfile & "\Downloads\"& xclbook &".xlsx")


Set objworksheet = objWorkbook.worksheets("Sheet1")
objworksheet.Range("A:G").AutoFilter 6,Array("Locked"),7,True
objWorkbook.ActiveSheet.AutoFilter.range.copy
objworkbook.sheets.Add
objWorkbook.ActiveSheet.paste
objWorkbook.worksheets("Sheet1").activate
objworksheet.Range("A:G").AutoFilter 6,Array("Missing","Rejected","Submitted","Not Submitted","Approved"),7,True
objWorkbook.ActiveSheet.AutoFilter.range.copy
objworkbook.sheets.Add
objWorkbook.ActiveSheet.paste
objWorkbook.worksheets("Sheet3").activate
objWorkbook.worksheets("Sheet3").Range("H2:H2").Formula = "=CONCATENATE(A2,D2,E2)"
usedrange = objWorkbook.worksheets("Sheet3").usedrange.rows.count
objWorkbook.worksheets("Sheet3").Range("H2:H2").copy
objWorkbook.worksheets("Sheet3").range("H3:H"&usedrange).pastespecial(-4104)
objWorkbook.worksheets("Sheet2").activate
objWorkbook.worksheets("Sheet2").Range("H2:H2").Formula = "=CONCATENATE(A2,D2,E2)"
usedrange1 = objWorkbook.worksheets("Sheet2").usedrange.rows.count
objWorkbook.worksheets("Sheet2").Range("H2:H2").copy
objWorkbook.worksheets("Sheet2").range("H3:H"&usedrange1).pastespecial(-4104)
objWorkbook.worksheets("Sheet3").activate
objWorkbook.worksheets("Sheet3").Range("I2:I2").Formula = "=Vlookup(H2,Sheet2!H:H,1,False)"
objWorkbook.worksheets("Sheet3").Range("I2:I2").copy
objWorkbook.worksheets("Sheet3").range("I3:I"&usedrange).pastespecial(-4104)
objworkbook.worksheets("Sheet3").range("A:I").AutoFilter 9,Array("#N/A"),7,True
objWorkbook.ActiveSheet.AutoFilter.range.copy
objworkbook.sheets.Add
objWorkbook.ActiveSheet.paste
objworkbook.worksheets("Sheet4").range("I:I").delete
objworkbook.worksheets("Sheet4").range("H:H").delete
objExcel.DisplayAlerts = False
objworkbook.worksheets("Sheet3").delete
objworkbook.worksheets("Sheet2").delete
objworkbook.worksheets("Sheet1").delete
objExcel.DisplayAlerts = True
objworkbook.worksheets("Sheet4").name = "Sheet1"
objWorkbook.save
objWorkbook.close
