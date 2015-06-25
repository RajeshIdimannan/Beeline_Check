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


objWorkbook.save
objWorkbook.close
