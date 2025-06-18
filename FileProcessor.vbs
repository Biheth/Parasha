Set args = WScript.Arguments

If args.Count < 3 Then
    MsgBox "������������ ����������!" & vbCrLf & _
           "�������������: FileProcessor.vbs [xls] [docx] [����]", vbCritical
    WScript.Quit
End If

xlsPath = args(0)
docxPath = args(1)
selectedDate = args(2)

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(xlsPath) Then
    MsgBox "���� XLSX �� ������:" & vbCrLf & xlsPath, vbCritical
    WScript.Quit
End If

If Not fso.FileExists(docxPath) Then
    MsgBox "���� DOCX �� ������:" & vbCrLf & docxPath, vbCritical
    WScript.Quit
End If

' �������������� ����
dateParts = Split(selectedDate, ".")
d = CInt(dateParts(0))
m = CInt(dateParts(1))
y = CInt(dateParts(2))
processedDate = DateSerial(y, m, d)

MsgBox "��������� ������ ������:" & vbCrLf & _
       "XLS: " & xlsPath & vbCrLf & _
       "DOCX: " & docxPath & vbCrLf & _
       "����: " & FormatDateTime(processedDate, 1), vbInformation, "��������� ���������"

' ���� ������ ��������� ������...