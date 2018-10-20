Option Explicit
' https://superuser.com/questions/641471/how-can-i-automatically-convert-powerpoint-to-pdf

Sub WriteLine ( strLine )
    WScript.Stdout.WriteLine strLine
End Sub

' http://msdn.microsoft.com/en-us/library/office/aa432714(v=office.12).aspx
Const msoFalse = 0   ' False.
Const msoTrue = -1   ' True.

' https://msdn.microsoft.com/de-de/vba/word-vba/articles/wdexportoptimizefor-enumeration-word
Const wdExportOptimizeForOnScreen = 1
Const wdExportOptimizeForPrint = 0

' https://msdn.microsoft.com/de-de/vba/word-vba/articles/wdexportformat-enumeration-word
Const wdExportFormatPDF = 17  ' XPS format

' https://msdn.microsoft.com/de-de/vba/word-vba/articles/wdexportrange-enumeration-word
Const wdExportAllDocument = 0

'
' This is the actual script
'

Dim inputFile
Dim outputFile
Dim objWRD
Dim objDocument
Dim objPrintOptions
Dim objFso
Dim doCloseOnEnd

If WScript.Arguments.Count <> 2 Then
    WriteLine "You need to specify input and output files."
    WScript.Quit
End If

inputFile = WScript.Arguments(0)
outputFile = WScript.Arguments(1)

Set objFso = CreateObject("Scripting.FileSystemObject")

If Not objFso.FileExists( inputFile ) Then
    WriteLine "Unable to find your input file " & inputFile
    WScript.Quit
End If

If objFso.FileExists( outputFile ) Then
    WriteLine "Your output file (' & outputFile & ') already exists!"
    WScript.Quit
End If

WriteLine "Input File:  " & inputFile
WriteLine "Output File: " & outputFile

On Error Resume Next
' Try to grab a running instance of Word...
Set objWRD = GetObject(, "Word.Application")

' Did we find anything...?
If Not TypeName(objWRD) = "Empty" Then
    'PowerPoint is already Running. Do not close PPT after script finished
	doCloseOnEnd = False
Else
    ' PowerPoitn is not running. Start it and close windows after script finished
	Set objWRD = CreateObject( "Word.Application" )
	doCloseOnEnd = True
End If
On Error Goto 0


objWRD.Visible = True
objWRD.Documents.Open inputFile

Set objDocument = objWRD.ActiveDocument

' Reference for this at http://msdn.microsoft.com/en-us/library/office/ff746080.aspx
objDocument.ExportAsFixedFormat outputFile, wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportAllDocument

objDocument.Close
If doCloseOnEnd Then
	objWRD.Quit
End If
