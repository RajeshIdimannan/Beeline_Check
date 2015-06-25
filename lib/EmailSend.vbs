Set objShell = CreateObject("Wscript.Shell")

Set oWS = WScript.CreateObject("WScript.Shell")

opfolder = WScript.Arguments.Item(0)
'msgbox WScript.Arguments.Item(0)
strPath = objShell.CurrentDirectory
'userProfile = oWS.ExpandEnvironmentStrings( "%userprofile%" )  
'Msgbox strPath
Dim filesys
Set filesys = CreateObject("Scripting.FileSystemObject")
'varPathCurrent = filesys.GetParentFolderName(WScript.ScriptFullName)
'varPathParent = filesys.GetParentFolderName(varPathCurrent)

 
'Msgbox varPathParent 
Dim outobj, mailobj
          Dim strBodyText
          Dim objFileToRead

          Set outobj = CreateObject("Outlook.Application")
          Set mailobj = outobj.CreateItem(0)
	
          strBodyText ="Hi Rajesh,"&vbNewLine&vbNewLine&"Please find attached beeline daily status report Sheet."&vbNewLine&"Please revert back to kishankumar.panda2@target.com for any queries."&vbNewLine&vbNewLine&"Thanks and Regards"&vbNewLine&"Kishan"

            With mailobj
'for multiple addresses please add ; in middle of mail id's
            .To = "kishankumar.panda2@target.com"
            .Subject = "Daily Status:Beeline_Report(Jan-Feb)"
            .Body = strBodyText
	.Attachments.Add(opfolder & "\BeelineReport.xlsx")
	
 	
 	.cc = "kishankumar.panda2@target.com"
	 
            .Display
          End With
	mailobj.Send
          'Clear the memory
          Set outobj = Nothing
          Set mailobj = Nothing

    

       