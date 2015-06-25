Set objShell = CreateObject("Wscript.Shell")
strPath = objShell.CurrentDirectory
'userProfile = oWS.ExpandEnvironmentStrings( "%userprofile%" )  
'Msgbox strPath
Dim filesys
Set filesys = CreateObject("Scripting.FileSystemObject")
varPathCurrent = filesys.GetParentFolderName(WScript.ScriptFullName)
varPathParent = filesys.GetParentFolderName(varPathCurrent)

 
'Msgbox varPathParent 
Dim outobj, mailobj
          Dim strBodyText
          Dim objFileToRead

          Set outobj = CreateObject("Outlook.Application")
          Set mailobj = outobj.CreateItem(0)
          strBodyText ="Hi All,"&vbNewLine&vbNewLine&"Please find attached beeline report Sheets"&vbNewLine&vbNewLine&"Thanks and Regards"

            With mailobj
            .To = "kishankumar.panda@target.com"
            .Subject = "Beeline_Automation_Report"
            .Body = strBodyText
	.Attachments.Add(varPathParent &"\BeelineReport@27-11-2014\BeelineReport.xlsx")
	
 	
 	.cc = "kishankumar.panda@target.com"
	 
            .Display
          End With
	mailobj.Send
          'Clear the memory
          Set outobj = Nothing
          Set mailobj = Nothing

    

       