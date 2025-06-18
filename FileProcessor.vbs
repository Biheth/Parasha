Set args = WScript.Arguments

If args.Count < 3 Then
    MsgBox "Недостаточно параметров!" & vbCrLf & _
           "Использование: FileProcessor.vbs [xls] [docx] [дата]", vbCritical
    WScript.Quit
End If

xlsPath = args(0)
docxPath = args(1)
selectedDate = args(2)

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(xlsPath) Then
    MsgBox "Файл XLSX не найден:" & vbCrLf & xlsPath, vbCritical
    WScript.Quit
End If

If Not fso.FileExists(docxPath) Then
    MsgBox "Файл DOCX не найден:" & vbCrLf & docxPath, vbCritical
    WScript.Quit
End If

' Преобразование даты
dateParts = Split(selectedDate, ".")
d = CInt(dateParts(0))
m = CInt(dateParts(1))
y = CInt(dateParts(2))
processedDate = DateSerial(y, m, d)

MsgBox "Обработка файлов начата:" & vbCrLf & _
       "XLS: " & xlsPath & vbCrLf & _
       "DOCX: " & docxPath & vbCrLf & _
       "Дата: " & FormatDateTime(processedDate, 1), vbInformation, "Параметры обработки"

' Ваша логика обработки файлов...