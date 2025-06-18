Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptPath = WScript.ScriptFullName
scriptFolder = fso.GetParentFolderName(scriptPath)
htaPath = scriptFolder & "\FileSelector.hta"
tempFile = scriptFolder & "\SelectedFiles.txt"

If fso.FileExists(tempFile) Then fso.DeleteFile tempFile

shell.Run """" & htaPath & """", 1, True

If fso.FileExists(tempFile) Then
    Set file = fso.OpenTextFile(tempFile, 1)
    xlsPath = file.ReadLine
    docxPath = file.ReadLine
    selectedDate = file.ReadLine ' ������ ����
    file.Close
    fso.DeleteFile tempFile

    MsgBox "��������� ���������:" & vbCrLf & _
           "XLS: " & xlsPath & vbCrLf & _
           "DOCX: " & docxPath & vbCrLf & _
           "����: " & selectedDate, vbInformation, "������ ��������"
    
    processorScript = scriptFolder & "\FileProcessor.vbs"
    If fso.FileExists(processorScript) Then
        ' �������� ��� ���������: xlsPath, docxPath, selectedDate
        command = "wscript """ & processorScript & """ """ & xlsPath & """ """ & docxPath & """ """ & selectedDate & """"
        shell.Run command, 1, False
    Else
        MsgBox "������ ��������� �� ������!", vbExclamation
    End If
Else
    MsgBox "�������� ��������", vbExclamation
End If