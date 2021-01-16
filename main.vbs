Option Explicit

Sub RemoveWhiteSpaces()
    Dim row As Long
    Dim pos As Integer
    Dim currentCol As String
    
    'the following statement find the position of the next $ characher.
    pos = InStr(2, ActiveCell.Address, "$")
    
    currentCol = Mid(ActiveCell.Address, 2, pos - 2)
    
    If currentCol = "" Then
        MsgBox "Please select a cell in which you want to filter out its column"
        Exit Sub
    End If
      
    For row = 1 To 100000
         Range(currentCol & row).Select
         If ActiveCell.Text = "" Then
            MsgBox "Finished"
            Exit Sub
         End If
         ActiveCell.FormulaR1C1 = RemoveSpecialCharacters(ActiveCell.FormulaR1C1)
    Next row
    
End Sub

Private Function RemoveSpecialCharacters(data As String)
    Dim i As Integer
    Dim statement As String
    Dim letter As Integer
    
    For i = 1 To Len(data)
        letter = Asc(Mid(data, i, 1))
        
        'the following if statement only check the valid characters to collect them.
        If letter = 32 Or letter = 170 Or (letter >= 192 And letter <= 214) Or (letter >= 216 And letter <= 219) Or (letter >= 221 And letter <= 223) Or letter = 225 Or (letter >= 227 And letter <= 230) Or letter = 236 Or letter = 237 Then
            statement = statement + Chr(letter)
        End If
    Next i
    
    RemoveSpecialCharacters = TrimExtraSpaces(Trim(statement))
End Function

Private Function TrimExtraSpaces(data As String)
    Dim parts() As String
    Dim i As Integer
    Dim statement As String
    
    data = Trim(data)
    
    'The following statment is added by me to check whether the passing string is empty (exit immediatly) or not.
    If data = "" Then
        Exit Function
    End If
    
    parts = SplitRemoveEmptyEntries(data, " ")
    
    For i = 0 To UBound(parts)
        If parts(i) <> " " Or parts(i) <> "" Then
            statement = statement + Trim(parts(i)) + " "
        End If
    Next i
    
    TrimExtraSpaces = Trim(statement)
End Function

'This function from Stackoverflow website
'https://stackoverflow.com/questions/37952387/vba-string-split-remove-empty-entries
Private Function SplitRemoveEmptyEntries(strInput As String, strDelimiter As String) As String()
    Dim strTmp As Variant
    Dim sSplit() As String
    Dim sSplitOut() As String
    ReDim Preserve sSplitOut(0)
    
    'The following statment is added by me to check whether the passing string is empty (exit immediatly) or not.
    If strInput = "" Then
        Exit Function
    End If
        
    For Each strTmp In Split(strInput, strDelimiter)
      If Trim(strTmp) <> "" Then
        ReDim Preserve sSplitOut(UBound(sSplitOut) + 1)
        sSplitOut(UBound(sSplitOut) - 1) = strTmp
      End If
    Next strTmp
    
    ReDim Preserve sSplitOut(UBound(sSplitOut) - 1)
    SplitRemoveEmptyEntries = sSplitOut
    
End Function

