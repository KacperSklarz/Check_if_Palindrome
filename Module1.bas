Attribute VB_Name = "Module1"
Sub Check()

    Range("C1:F11").Value = " "
    Range("C1") = "Iteration"
    Range("D1") = "Value"
    Range("E1") = "Binary System"
    Range("F1") = "Squared Value"
    
    Dim n As String
    n = Range("B1").Value
    
    If Not IsNumeric(n) Or Int(n) <> n Then
        MsgBox ("Niepoprawne dane")
        Exit Sub
    End If
    
    If (n < 1 Or n > 999999999) Then
        MsgBox ("Dane Powinny byæ 1<n<999999999")
        Exit Sub
    End If
    
    
   Dim x As Long
   x = n
   
   Dim i As Integer
   For i = 0 To 9
        
        Dim x_ As Long, bin As String, square As String
        x_ = x
        bin = ""
        
        While x_ > 0
            bin = CStr(x_ Mod 2) & bin
            x_ = x_ \ 2
        Wend
     
        square = (bin) ^ 2
        
        Range("C2").Offset(i, 0).Value = "Iteration" & i
        Range("D2").Offset(i, 0).Value = x
        Range("E2").Offset(i, 0).Value = bin
        Range("F2").Offset(i, 0).Value = square
        kwadrat = ""
        dwa = ""
        bin = ""
        
        Dim len_ As Integer
        len_ = Len(Str(x)) - 1
        
        If len_ = 1 Then
            MsgBox ("Liczba" & n & "jest palindromem")
            Exit Sub
        End If
        
        x_ = x
        Dim j As Integer
        
        
        Dim temp_1, reversed As Long
        reversed = 0
        
        While x_ <> 0
            temp_1 = x_ Mod 10
            reversed = reversed * 10 + temp_1
            x_ = (x_ - temp_1) / 10
        Wend
        
        
        If i = 9 And x <> reversed Then
            MsgBox ("Mimo Iteracji liczba nie jest palindromem")
        End If
        
        If x = reversed And i = 0 Then
            MsgBox ("Pocz¹tkowa liczba " & n & " jest palindromem")
            Exit Sub
        ElseIf x = reversed Then
            MsgBox ("Liczba" & n & " jest palindromem, potrzeba by³o " & i & " iteracji")
            Exit Sub
        End If
        
        
    
        Dim suma As Long
        sum_ = x + reversed
        x = sum_
        sum_ = 0
        reversed = 0
        temp_1 = 0
    Next i
    
    
    
    
End Sub
