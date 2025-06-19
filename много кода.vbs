' Функция проверки наличия греческих букв
Function HasGreek(str)
    Dim greek, i
    greek = "αβγδεζηθικλμνξοπρστυφχψω"
    HasGreek = False
    For i = 1 To Len(greek)
        If InStr(str, Mid(greek, i, 1)) > 0 Then
            HasGreek = True
            Exit Function
        End If
    Next
End Function

' Функция нормализации названия фидера
Function NormalizeFeeder(feederStr)
    Dim parts, i, j, temp
    parts = Split(Replace(feederStr, " ", ""), "+")
    For i = 0 To UBound(parts)
        For j = i + 1 To UBound(parts)
            If parts(i) > parts(j) Then
                temp = parts(i)
                parts(i) = parts(j)
                parts(j) = temp
            End If
        Next
    Next
    NormalizeFeeder = Join(parts, "+")
End Function

' Функция извлечения кандидатов в фидеры
Function ExtractFeederCandidates(line)
    Dim validChars, greekLetters, candidates(), current, char, i, startPos, inCandidate, j, c
    validChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюяαβγδεζηθικλμνξοπρστυφχψω-+ "
    greekLetters = "αβγδεζηθικλμνξοπρστυφχψω"
    
    ReDim candidates(-1)
    current = ""
    startPos = 0
    inCandidate = False
    
    For i = 1 To Len(line)
        char = Mid(line, i, 1)
        If InStr(validChars, char) > 0 Then
            If Not inCandidate Then
                inCandidate = True
                startPos = i - 1
            End If
            current = current & char
        Else
            If inCandidate Then
                If Len(current) >= 2 Then
                    Dim hasDigit, hasGreek
                    hasDigit = False
                    hasGreek = False
                    For j = 1 To Len(current)
                        c = Mid(current, j, 1)
                        If InStr("0123456789", c) > 0 Then hasDigit = True
                        If InStr(greekLetters, c) > 0 Then hasGreek = True
                    Next
                    If hasDigit Or hasGreek Then
                        ReDim Preserve candidates(UBound(candidates) + 1)
                        candidates(UBound(candidates)) = Array(current, startPos, Len(current))
                    End If
                End If
                current = ""
                inCandidate = False
            End If
        End If
    Next

    If inCandidate And Len(current) >= 2 Then
        hasDigit = False
        hasGreek = False
        For j = 1 To Len(current)
            c = Mid(current, j, 1)
            If InStr("0123456789", c) > 0 Then hasDigit = True
            If InStr(greekLetters, c) > 0 Then hasGreek = True
        Next
        If hasDigit Or hasGreek Then
            ReDim Preserve candidates(UBound(candidates) + 1)
            candidates(UBound(candidates)) = Array(current, startPos, Len(current))
        End If
    End If
    
    ExtractFeederCandidates = candidates
End Function

' Основная функция замены фидера
Function ReplaceFeeder(line, oldFeeder, newFeeder)
    Dim candidates, withGreek(), withoutGreek(), i, candidate, startPos, length
    Dim midLine, minDist, centerP, dist, target, normCandidate, normOld
    ReDim withGreek(-1)
    ReDim withoutGreek(-1)
    
    candidates = ExtractFeederCandidates(line)
    If Not IsArray(candidates) Then
        ReplaceFeeder = line
        Exit Function
    End If
    
    For i = 0 To UBound(candidates)
        candidate = candidates(i)(0)
        startPos = candidates(i)(1)
        length = candidates(i)(2)
        
        If HasGreek(candidate) Then
            ReDim Preserve withGreek(UBound(withGreek) + 1)
            withGreek(UBound(withGreek)) = Array(candidate, startPos, length)
        Else
            ReDim Preserve withoutGreek(UBound(withoutGreek) + 1)
            withoutGreek(UBound(withoutGreek)) = Array(candidate, startPos, length)
        End If
    Next
    
    target = Array("", -1, 0)
    midLine = Len(line) / 2
    
    If UBound(withGreek) >= 0 Then
        minDist = 1E+38
        For i = 0 To UBound(withGreek)
            startPos = withGreek(i)(1)
            length = withGreek(i)(2)
            centerP = startPos + (length / 2)
            dist = Abs(centerP - midLine)
            If dist < minDist Then
                minDist = dist
                target = withGreek(i)
            End If
        Next
    ElseIf UBound(withoutGreek) >= 0 Then
        minDist = 1E+38
        For i = 0 To UBound(withoutGreek)
            startPos = withoutGreek(i)(1)
            length = withoutGreek(i)(2)
            centerP = startPos + (length / 2)
            dist = Abs(centerP - midLine)
            If dist < minDist Then
                minDist = dist
                target = withoutGreek(i)
            End If
        Next
    Else
        ReplaceFeeder = line
        Exit Function
    End If
    
    normCandidate = NormalizeFeeder(target(0))
    normOld = NormalizeFeeder(oldFeeder)
    
    If normCandidate = normOld Then
        ReplaceFeeder = Left(line, target(1)) & newFeeder & Mid(line, target(1) + target(2) + 1)
    Else
        ReplaceFeeder = line
    End If
End Function

' ==================== ТЕСТИРОВАНИЕ ====================
Dim testLines, oldFeeder, newFeeder, line, result
testLines = Array( _
    "прпр 12вр выори. 123549α+γ c dhjb 32342 skj", _
    "ghuyk 04 dh. fg 23 β ckjh shdf 4787 f", _
    "dkh ПРР-8 +  УНГ от ТР2 +1234 δ (ывлда 555) выармс -вл.", _
    "Резерв 555α основной 123β+γ и 789δ-ε", _
    "Авария на 15А (резерв) и 42β+γ основной" _
)

oldFeeder = "123549α+γ"
newFeeder = "Ф-551А+Б"

For Each line In testLines
    result = ReplaceFeeder(line, oldFeeder, newFeeder)
    WScript.Echo "Исходная: " & line
    WScript.Echo "Результат: " & result
    WScript.Echo "----------------------------------------"
Next
