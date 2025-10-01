' Module: Normalization.bas
' Author: Mohammed Saqlain Altaf
' Tool: Excel Normalization Tool
' Version: 1.0.0
' Date: 2025-10-01
' License: MIT
' Description: Normalizes text for AI training (acronyms, numbers, expansions).

Option Explicit

' ==== MAIN ENTRY POINTS ====
Public Sub NormalizeSelectionUpper()
    Call NormalizeSelection(True)
End Sub

Public Sub NormalizeSelectionLower()
    Call NormalizeSelection(False)
End Sub

Private Sub NormalizeSelection(ByVal toUpper As Boolean)
    Dim cell As Range
    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            cell.Value = NormalizeText(CStr(cell.Value), toUpper)
        End If
    Next cell
End Sub

' ==== CORE NORMALIZER ====
Private Function NormalizeText(ByVal txt As String, ByVal toUpper As Boolean) As String
    Dim words() As String, result As String, i As Long, w As String
    
    txt = RemovePunctuation(txt)
    words = Split(txt, " ")
    
    For i = LBound(words) To UBound(words)
        w = Trim(words(i))
        If w = "" Then GoTo SkipWord
        
        ' Word expansions (MR, DR, etc.)
        If ExpansionDict.Exists(UCase(w)) Then
            w = ExpansionDict(UCase(w))
        
        ' Acronyms
        ElseIf IsAllCaps(w) And Len(w) <= 5 Then
            If AcronymExceptions.Exists(UCase(w)) Then
                w = UCase(w) ' keep as word
            Else
                w = SeparateLetters(w) ' spell out
            End If
        
        ' Numbers
        ElseIf IsNumeric(w) Then
            w = NormalizeNumberOrYear(CLng(w))
        
        Else
            w = UCase(w)
        End If
        
SkipWord:
        If result = "" Then
            result = w
        Else
            result = result & " " & w
        End If
    Next i
    
    result = Replace(result, "  ", " ")
    result = Trim(result)
    
    If toUpper Then
        NormalizeText = UCase(result)
    Else
        NormalizeText = LCase(result)
    End If
End Function

' ==== NUMBER HANDLER (with year detection) ====
Private Function NormalizeNumberOrYear(ByVal num As Long) As String
    If num >= 1000 And num <= 2100 Then
        NormalizeNumberOrYear = YearToWords(num)
    Else
        NormalizeNumberOrYear = NumberToWords(num)
    End If
End Function

Private Function YearToWords(ByVal yr As Long) As String
    Dim res As String
    
    If yr >= 2000 And yr <= 2009 Then
        res = "two thousand"
        If yr Mod 100 > 0 Then res = res & " " & NumberToWords(yr Mod 100)
    
    ElseIf yr >= 2010 And yr <= 2099 Then
        res = "twenty " & NumberToWords(yr Mod 100)
    
    ElseIf yr >= 1900 And yr <= 1999 Then
        res = NumberToWords(yr \ 100) & " " & NumberToWords(yr Mod 100)
    
    Else
        res = NumberToWords(yr)
    End If
    
    YearToWords = Trim(res)
End Function

' ==== WORD EXPANSION DICTIONARY ====
Private Function ExpansionDict() As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    dict.Add "DR", "DOCTOR"
    dict.Add "MR", "MISTER"
    dict.Add "MRS", "MISSUS"
    dict.Add "VS", "VERSUS"
    dict.Add "ST", "SAINT"
    dict.Add "PROF", "PROFESSOR"
    dict.Add "JR", "JUNIOR"
    Set ExpansionDict = dict
End Function

' ==== ACRONYM EXCEPTIONS (kept as a word) ====
Private Function AcronymExceptions() As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    dict.Add "NASA", True
    dict.Add "FIFA", True
    dict.Add "NEET", True
    dict.Add "UNESCO", True
    Set AcronymExceptions = dict
End Function

' ==== HELPERS ====
Private Function RemovePunctuation(ByVal txt As String) As String
    Dim i As Long, c As String, res As String
    For i = 1 To Len(txt)
        c = Mid(txt, i, 1)
        If (c Like "[A-Za-z0-9]") Or c = " " Or c = "'" Then
            res = res & c
        End If
    Next i
    RemovePunctuation = res
End Function

Private Function IsAllCaps(ByVal w As String) As Boolean
    IsAllCaps = (UCase(w) = w And w Like "*[A-Z]*")
End Function

Private Function SeparateLetters(ByVal w As String) As String
    Dim i As Long, res As String
    For i = 1 To Len(w)
        res = res & Mid(w, i, 1) & " "
    Next i
    SeparateLetters = Trim(res)
End Function

' ==== NUMBERS TO WORDS ====
Private Function NumberToWords(ByVal num As Long) As String
    Dim Units As Variant, Tens As Variant
    Units = Array("", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", _
                  "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen")
    Tens = Array("", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")
    
    If num = 0 Then
        NumberToWords = "zero"
        Exit Function
    End If
    
    Dim res As String
    res = ""
    
    If num \ 1000 > 0 Then
        res = res & NumberToWords(num \ 1000) & " thousand "
        num = num Mod 1000
    End If
    If num \ 100 > 0 Then
        res = res & NumberToWords(num \ 100) & " hundred "
        num = num Mod 100
        If num > 0 Then res = res & "and "
    End If
    If num > 0 Then
        If num < 20 Then
            res = res & Units(num)
        Else
            res = res & Tens(num \ 10)
            If (num Mod 10) > 0 Then
                res = res & " " & Units(num Mod 10)
            End If
        End If
    End If
    
    NumberToWords = Trim(res)
End Function

