Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

Dim sOld, sNew
sOld = "e:\Rajkumar\PDFExcelExtract "
sNew = "e:\Rajkumar\PDFExcelExtract"

' Step 1 – rename folder to remove trailing space (safe to re-run)
If oFSO.FolderExists(sOld) And Not oFSO.FolderExists(sNew) Then
    Dim oFolder
    Set oFolder = oFSO.GetFolder(sOld)
    oFolder.Name = "PDFExcelExtract"
End If

' Step 2 – run the bat from the (now renamed) folder
Dim sBat
If oFSO.FolderExists(sNew) Then
    sBat = sNew & "\run_pdf.bat"
Else
    sBat = sOld & "\run_pdf.bat"
End If

oShell.Run "cmd.exe /k """ & sBat & """", 1, False
