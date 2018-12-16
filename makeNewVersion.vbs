' Make new version of the document

WorkingPath = "."

FilePrefix  = "DSP_PR00024880_5G-Radio_"
FileExt     = ".docx"
FileRegExp  = "^" & FilePrefix & "\d{8}" & FileExt & "$"

NewVersion  = Year(Date) & Right("0"& Month(Date), 2) & Right("0"& Day(Date), 2)
NewFileName = FilePrefix & NewVersion & FileExt
NewDate     = Right("0"& Day(Date), 2) & "-" & Right("0"& Month(Date), 2) & "-" & Year(Date)


Set fso = WScript.CreateObject("Scripting.FileSystemObject")
absWorkingPath = fso.GetAbsolutePathName(WorkingPath)

Function GetLastVersionFileName(aFileRegExp, aPath)
	Set regExp = WScript.CreateObject("VBScript.RegExp")
	regExp.Pattern = aFileRegExp
	Set fileNames = WScript.CreateObject("System.Collections.ArrayList")
	For Each fsoFile In fso.GetFolder(aPath).Files
		If regExp.Test(fsoFile.Name) Then
			fileNames.Add fsoFile.Name
		End If
	Next

	fileNames.Sort()

	If fileNames.Count = 0 Then
		GetLastVersionFileName = ""
	Else
		GetLastVersionFileName = fileNames.Item( fileNames.Count -1)
	End If
End Function

lastFileName = GetLastVersionFileName(FileRegExp, absWorkingPath)
Select Case lastFileName
Case ""
	MsgBox "No one version has found!"
	WScript.Quit 1
Case NewFileName
	MsgBox "New version already exists!"
	WScript.Quit 2
End Select

fso.CopyFile lastFileName, NewFileName

Set word = WScript.CreateObject("Word.Application")
word.Visible = True
Set doc = word.Documents.Open( fso.BuildPath(absWorkingPath, NewFileName) )
With doc
	.AcceptAllRevisions

	'' A Dependency
	.CustomDocumentProperties("Versionsnummer").Value = NewVersion
	.CustomDocumentProperties("Veröffentlichungsdatum").Value = NewDate

	.Fields.Update
	For Each section In .Sections
		For Each header In section.Headers
			header.Range.Fields.Update
		Next
		For Each footer In section.Footers
			footer.Range.Fields.Update
		Next
	Next

	.Save
End With

word.Quit

MsgBox "New version has created!"
WScript.Quit 0
