Attribute VB_Name = "Module1"
Option Explicit

Function Validate(TxB As TextBox) As Boolean
    On Error Resume Next
    Dim xHour As String, xMinute As String
    Dim xDay As String, xMonth As String, xYear As String
    Dim CtrlVal As String
    Dim xInput As String
    Dim xRepeat As Single
    Dim Crka As String
    Dim xNo As Single
    Dim MaxNo As Single, MinNo As Single
    Dim displayMsg As String
    
    CtrlVal = UCase$(TxB.Tag)
    
    If InStr(CtrlVal, "NOTEMPTY;") Then
        xInput = TxB.Text
        If xInput <> Empty Then
            Validate = True
        Else
            Validate = False
            GoSub displayMsg
            Exit Function
        End If
    End If
    
    If InStr(CtrlVal, "UCASE;") Then
        xInput = TxB.Text
        TxB.Text = UCase$(xInput)
        Validate = True
    
    ElseIf InStr(CtrlVal, "LCASE;") Then
        xInput = TxB.Text
        TxB.Text = LCase$(xInput)
        Validate = True
    End If
    
    If InStr(CtrlVal, "TIME;") Then
        xInput = TxB.Text
        'preglej èe je v xInputu katerikoli drugi znak kot : in èe je ga zamenjaj z :
        For xRepeat = 1 To Len(xInput)
            Crka = Mid$(xInput, xRepeat, 1)
            If Not IsNumeric(Crka) Then
                Mid$(xInput, xRepeat) = ":"
            End If
        Next
        TxB = xInput
        If Len(xInput) <= 2 Then
            TxB = TxB + ":0"
            xInput = TxB.Text
        ElseIf Len(xInput) = 3 Then
            xHour = Mid$(xInput, 1, 1)
            xMinute = Mid$(xInput, 2)
            TxB = xHour + ":" + xMinute
            xInput = TxB.Text
        ElseIf Len(xInput) = 4 Then
            xHour = Mid$(xInput, 1, 2)
            xMinute = Mid$(xInput, 3)
            TxB = xHour + ":" + xMinute
            xInput = TxB.Text
        End If
        xHour = Format$(xInput, "hh")
        xMinute = Format$(xInput, "nn")
        If Not IsNumeric(xHour) Then
            xHour = "HH"
        End If
        If Not IsNumeric(xMinute) Then
            xMinute = "MM"
        End If
        TxB = xHour + ":" + xMinute
        If Len(TxB) > 5 Then
            TxB = Left$(TxB, 5)
        End If
        If xHour = "HH" And xMinute = "MM" Then
            Validate = False
            GoSub displayMsg
            Exit Function
        Else
            Validate = True
        End If
    
    ElseIf InStr(CtrlVal, "DATE;") Then
        xInput = TxB.Text
        If UCase$(xInput) = "N" Then
            TxB = Format$(Now, "dd-mm-yyyy")
        End If
        If UCase$(xInput) = "T" Then
            TxB = Format$(Now + 1, "dd-mm-yyyy")
        End If
        If UCase$(xInput) = "A" Then
            TxB = Format$(Now + 2, "dd-mm-yyyy")
        End If
        If UCase$(xInput) = "Y" Then
            TxB = Format$(Now - 1, "dd-mm-yyyy")
        End If
        
        If Left$(xInput, 1) = "+" And Len(xInput) > 1 Then
            xNo = Val(Mid$(xInput, 2))
            TxB = Format$(Now + xNo, "dd-mm-yyyy")
        End If
        
        If Left$(xInput, 1) = "-" And Len(xInput) > 1 Then
            xNo = Val(Mid$(xInput, 2))
            TxB = Format$(Now - xNo, "dd-mm-yyyy")
        End If
        
        xInput = TxB.Text
        
        If Len(xInput) <= 2 Then
            TxB = TxB + "-" + Format$(Month(Now), "00")
            xInput = TxB.Text
        ElseIf Len(xInput) = 3 And InStr(xInput, "-") = 0 Then
            xDay = Mid$(xInput, 1, 1)
            xMonth = Mid$(xInput, 2)
            TxB = xDay + "-" + xMonth
            xInput = TxB.Text
            'preveri èe je ta datum sploh možen!
            If Not IsDate(xInput) Or Year(xInput) < 1990 Then
                xInput = xDay + xMonth
                xDay = Mid$(xInput, 1, 2)
                xMonth = Mid$(xInput, 3)
                TxB = xDay + "-" + xMonth
                xInput = TxB.Text
            End If
        ElseIf Len(xInput) = 4 And InStr(xInput, "-") = 0 Then
            xDay = Mid$(xInput, 1, 2)
            xMonth = Mid$(xInput, 3)
            TxB = xDay + "-" + xMonth
            xInput = TxB.Text
        End If
        If IsDate(xInput) Then
            If DateDiff("yyyy", "01-01-1990", xInput) < 0 Then
                TxB = "dd-mm-yyyy"
                Validate = False
                GoSub displayMsg
                Exit Function
            Else
                TxB = Format$(xInput, "dd-mm-yyyy")
                Validate = True
            End If
        Else
            TxB = "dd-mm-yyyy"
            Validate = False
            GoSub displayMsg
            Exit Function
        End If
    End If
    If InStr(CtrlVal, "NUMERIC;") Then
        xInput = TxB.Text
        If IsNumeric(xInput) Then
            Validate = True
        Else
            Validate = False
            GoSub displayMsg
            Exit Function
        End If
    End If
    If InStr(CtrlVal, "MAX=") Then
        xInput = TxB.Text
        MaxNo = Val(Mid(CtrlVal, InStr(CtrlVal, "MAX=") + 4))
        If Val(xInput) <= MaxNo Then
            Validate = True
        Else
            Validate = False
            GoSub displayMsg
            Exit Function
        End If
    End If
    If InStr(CtrlVal, "MIN=") Then
        xInput = TxB.Text
        MinNo = Val(Mid(CtrlVal, InStr(CtrlVal, "MIN=") + 4))
        If Val(xInput) >= MinNo Then
            Validate = True
        Else
            Validate = False
            GoSub displayMsg
            Exit Function
        End If
    End If
    Validate = True
    Exit Function

displayMsg:
    If InStr(CtrlVal, "DISPLAY=") Then
        displayMsg = Mid$(CtrlVal, InStr(CtrlVal, "DISPLAY=") + 8)
        displayMsg = Left$(displayMsg, InStr(displayMsg, ";") - 1)
        MsgBox displayMsg, vbCritical
    End If
Return
End Function

Sub DisplayErrTbX(TbX As TextBox)
    Dim OldColor As Single
    Dim xRepeat As Single
    OldColor = TbX.BackColor
    For xRepeat = 1 To 1
        TbX.BackColor = vbRed
        TbX.Refresh
        Delay 0.1
        TbX.BackColor = OldColor
        TbX.Refresh
        Delay 0.1
    Next
    Beep
End Sub
Sub Delay(Sec)
    Dim x As Single
    x = Timer
    Do
    Loop Until Abs(x - Timer) > Sec
End Sub


