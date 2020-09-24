VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "KeyLogger Log File Viewer"
   ClientHeight    =   7575
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmd 
      Left            =   3990
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstLog 
      Height          =   765
      Left            =   1320
      TabIndex        =   0
      Top             =   870
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1349
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Log Data"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Transformed Data"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Handle"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Class Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Window Text"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Log File"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close Log File"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSaveToFile 
         Caption         =   "Save To File"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type LogDB
    TimeStamp As Date
    Data As String
    Handle As Long
    ClassName As String
    WindowsText As String
End Type


Private Sub Form_Resize()
    If ScaleWidth < 100 Or ScaleHeight < 100 Then Exit Sub
    lstLog.Move 30, 30, ScaleWidth - 60, ScaleHeight - 90
End Sub

Private Sub lstLog_DblClick()
On Error Resume Next
    MsgBox lstLog.SelectedItem.SubItems(1)
End Sub

Private Sub lstLog_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstLog.ListItems.Count > 0 And Button = 2 Then PopupMenu mnuHidden
End Sub

Private Sub mnuFileClose_Click()
    lstLog.ListItems.Clear
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
Dim sFile As String

    With cmd
        .Filter = "KeyLogger Database (*.dat)|*.dat|All File(s)|*.*"
        .DialogTitle = "Select The Log File"
        .ShowOpen
        sFile = .FileName
    End With
    If Len(sFile) <= 0 Then Exit Sub
    
    LoadLog sFile, lstLog
    
End Sub

Private Function LoadLog(sFileName As String, lstView As ListView)
Dim db As LogDB
Dim fNum As Integer
Dim sTmp As String
    
    fNum = FreeFile
    
    Open sFileName For Binary As fNum
        lstView.ListItems.Clear
        While Not EOF(fNum)
            Get fNum, , db
            If Len(db.Data) > 0 Then
                sTmp = Dec(db.Data)
                lstView.ListItems.Add lstView.ListItems.Count + 1, , sTmp
                lstView.ListItems(lstView.ListItems.Count).ListSubItems.Add 1, , Transform(sTmp)
                lstView.ListItems(lstView.ListItems.Count).ListSubItems.Add 2, , db.Handle
                lstView.ListItems(lstView.ListItems.Count).ListSubItems.Add 3, , db.ClassName
                lstView.ListItems(lstView.ListItems.Count).ListSubItems.Add 4, , db.WindowsText
                
                
            End If
        Wend
    Close fNum

End Function

Private Function Transform(sData As String) As String
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

Private Sub mnuSaveToFile_Click()
Dim sFile As String
Dim nFile As Integer

    With cmd
        .Filter = "Text File|*.txt"
        .ShowSave
        sFile = .FileName
    End With
    nFile = FreeFile
    
    Open sFile For Output As nFile
        Print #nFile, "Data Captured: " & lstLog.SelectedItem.Text & vbCrLf
        Print #nFile, "Transformed Data: " & lstLog.SelectedItem.SubItems(1) & vbCrLf
    Close nFile

End Sub
