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
    selectedDate = file.ReadLine ' Читаем дату
    file.Close
    fso.DeleteFile tempFile

    MsgBox "Выбранные параметры:" & vbCrLf & _
           "XLS: " & xlsPath & vbCrLf & _
           "DOCX: " & docxPath & vbCrLf & _
           "Дата: " & selectedDate, vbInformation, "Данные получены"
    
    processorScript = scriptFolder & "\FileProcessor.vbs"
    If fso.FileExists(processorScript) Then
        ' Передаем три параметра: xlsPath, docxPath, selectedDate
        command = "wscript """ & processorScript & """ """ & xlsPath & """ """ & docxPath & """ """ & selectedDate & """"
        shell.Run command, 1, False
    Else
        MsgBox "Скрипт обработки не найден!", vbExclamation
    End If
Else
    MsgBox "Операция отменена", vbExclamation
End If