Attribute VB_Name = "XOR_EncDec"

Function Enc(ByVal sStr As String) As String

On Error GoTo Debugger

    Dim XOR_Array() As String
    Dim i As Integer
    Dim sTempData As String
    
    ReDim XOR_Array(Len(sStr))
    
    For i = 1 To Len(sStr)
            XOR_Array(i) = Chr(Int(Rnd * 255))
            sTempData = XOR_Array(i) & sTempData & Chr(Asc(Mid(sStr, i, 1)) Xor Asc(XOR_Array(i)))
    Next i
    
    Enc = sTempData
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case MsgBox(Error$, vbExclamation + vbAbortRetryIgnore + vbDefaultButton1, "Error In Decryption")
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Function


Function Dec(ByVal sStr As String) As String

On Error GoTo Debugger

    Dim XOR_Array() As String
    Dim Str_Array() As String
        
    Dim i As Integer
    Dim sTempData As String
    
    ReDim XOR_Array(Len(sStr) / 2)
    ReDim Str_Array(Len(sStr) / 2)
    
    If sStr = "" Then Exit Function
    
    
    For i = 0 To (Len(sStr) / 2) - 1
        XOR_Array(i) = Mid(sStr, (Len(sStr) / 2) - i, 1)
        Str_Array(i) = Mid(sStr, (Len(sStr) / 2) + i + 1, 1)
    Next i
    
    For i = 0 To UBound(XOR_Array) - 1
        sTempData = sTempData & Chr(Asc(Str_Array(i)) Xor Asc(XOR_Array(i)))
    Next i
    Dec = sTempData
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case MsgBox(Error$, vbExclamation + vbAbortRetryIgnore + vbDefaultButton1, "Error In Decryption")
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Function

