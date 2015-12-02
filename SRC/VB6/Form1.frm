VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2415
   ClientLeft      =   1710
   ClientTop       =   6975
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   9915
   Begin VB.Timer tmrPi 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4800
      Top             =   1080
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "txtStatus"
      Top             =   2040
      Width           =   9735
   End
   Begin VB.Timer tmrGateway 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4800
      Top             =   600
   End
   Begin VB.Timer tmrCMD 
      Interval        =   1500
      Left            =   4800
      Top             =   120
   End
   Begin VB.Frame frmPihole 
      Caption         =   "frmPihole"
      Height          =   1815
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton btnScanPi 
         Caption         =   "btnScanPi"
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblPiHole 
         Caption         =   "lblPiHole"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   4575
      End
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0442
      Top             =   2520
      Width           =   9735
   End
   Begin VB.Frame frmGW 
      Caption         =   "frmGW"
      Height          =   1815
      Left            =   100
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton btnScan 
         Caption         =   "btnScan"
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblGW 
         Caption         =   "lblGW"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpszOp As String, _
     ByVal lpszFile As String, ByVal lpszParams As String, _
     ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim WithEvents StdIO As cStdIO
Attribute StdIO.VB_VarHelpID = -1
Dim bExitAfterCancel As Boolean

Private Sub btnScan_Click()
On Error GoTo btnScan_Click_Error       'error handler

lblGW.Caption = "Unknown"
lblGW.FontStrikethru = True

Dim strGateway As String
Dim lBytesWritten As Long
txtOutput.Text = ""
txtStatus.Text = "PROCESSING GATEWAY"
btnScan.Enabled = False     'disable button to allow cmd to process data
strGateway = "ipconfig | find " & Chr(34) & "Default" & Chr(34)

lBytesWritten = StdIO.WriteData(strGateway)

tmrGateway.Enabled = True   'start up a wait timer timer to fire up gateway processing

   On Error GoTo 0
   Exit Sub
btnScan_Click_Error:        'error handler sub
'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnScan_Click of Form Form1"
End Sub

Private Sub btnScanPi_Click()
On Error GoTo btnScanPi_Click_Error     'error handler

lblPiHole.Caption = "Unknown"
lblPiHole.FontStrikethru = True

Dim strPiHole As String
Dim lBytesWritten2 As Long
txtOutput.Text = ""
txtStatus.Text = "PROCESSING PI-HOLE"
btnScanPi.Enabled = False   'disable pi scan button to allow cmd to process data
strPiHole = "arp -a | find " & Chr(34) & "b8-27-eb" & Chr(34)   'scan for Raspberry MACs like b8-27-eb

lBytesWritten2 = StdIO.WriteData(strPiHole)

tmrPi.Enabled = True   'start up a wait timer to fire up pi-hole processing

   On Error GoTo 0
   Exit Sub
btnScanPi_Click_Error:      'error handler sub
'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnScanPi_Click of Form Form1"
End Sub

Private Sub ProcessPi()
On Error GoTo ProcessPi_Error   'error handler

'btnScanPi.Enabled = True    'turn on button for another try
Dim strPiHole As String

strPiHole = "arp -a | find " & Chr(34) & "b8-27-eb" & Chr(34)
txtOutput.Text = Replace$(txtOutput.Text, strPiHole, "")   'removes command string from output
txtOutput.Text = Replace$(txtOutput.Text, App.Path & ">", "")     'last line from output

txtOutput.Text = Replace(txtOutput.Text, vbCrLf, "-")   'add delimiting chars to check for blanks

If txtOutput.Text = "--" Then         'check here if string is null. If null pi was not found so exit sub
lblPiHole.Caption = "NOT FOUND"
txtStatus.Text = "Pi is not found. Try a ping sweep to wake up Pi?"
btnScanPi.Enabled = True    'turn on button for anther try
Exit Sub
Else
End If

txtOutput.Text = Replace(txtOutput.Text, "-", vbCrLf)   'remove delimiting chars to continue
txtOutput.Text = Left$(txtOutput.Text, 21)  'get the first 21 characters of the textbox

txtOutput.Text = LineTrim(txtOutput.Text)   'removes blank lines
txtOutput.Text = Trim$(txtOutput.Text)      'removes leading and trailing spaces

btnScanPi.Enabled = True    'turn on button for anther try

txtStatus.Text = "Pi-Hole found at " & txtOutput.Text

lblPiHole.Caption = "http://" & txtOutput.Text & "/admin"
lblPiHole.FontStrikethru = False

   On Error GoTo 0
   Exit Sub
ProcessPi_Error:    'error handler sub
'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcessPi of Form Form1"
End Sub

Private Sub ProcessGateway()
On Error GoTo ProcessGateway_Error   'error handler

btnScan.Enabled = True      'turn on button for another try
Dim strGateway As String
strGateway = "ipconfig | find " & Chr(34) & "Default" & Chr(34)
txtOutput.Text = Replace$(txtOutput.Text, strGateway, "")   'removes command string from output
txtOutput.Text = Replace$(txtOutput.Text, App.Path & ">", "")   'last line from output
txtOutput.Text = Replace$(txtOutput.Text, "Default Gateway", "")   'removes default gateway text
txtOutput.Text = Replace$(txtOutput.Text, " . ", "")    'replace period/spaces with null
txtOutput.Text = Replace$(txtOutput.Text, ": ", "")     'replace colon with null
txtOutput.Text = Replace$(txtOutput.Text, "..", "")     'removes extra periods with null

'txtOutput.Text = Replace(txtOutput.Text, vbCrLf, "-")   'add delimiting chars to check for blanks

If txtOutput.Text = "--" Then         'check here if string is null. If null gateway was not found so exit sub
lblGW.Caption = "NOT FOUND"
txtStatus.Text = "Gateway was not found. Your network is not up."
btnScan.Enabled = True      'turn on button for another try
Exit Sub
Else
End If
'Exit Sub    'remove me later

txtOutput.Text = LineTrim(txtOutput.Text)   'removes blank lines
txtOutput.Text = Trim$(txtOutput.Text)      'removes leading and trailing spaces
btnScan.Enabled = True

txtStatus.Text = "Gateway found at " & txtOutput.Text

lblGW.Caption = "http://" & txtOutput.Text
lblGW.FontStrikethru = False

   On Error GoTo 0
   Exit Sub
ProcessGateway_Error:       'error handler sub
'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcessGateway of Form Form1"
End Sub

Function LineTrim(ByVal sText As String) As String
    '-- add pre Lf and post Cr
    sText = vbLf & sText & vbCr '<--- unusual characters added
    '-- remove all spaces at the end of lines
    Do While InStr(sText, " " & vbCr)
        sText = Replace(sText, " " & vbCr, vbCr)
    Loop
    '-- remove all multiple CrLf's
    sText = Replace(sText, vbLf & vbCr, "") '<--- westconn1's smart line
    '-- remove first and last added characters
    LineTrim = Mid$(sText, 2, Len(sText) - 2)
End Function

Private Sub Form_Load()
Set StdIO = New cStdIO

'Set up app labels and buttons
Form1.Caption = "Pi-Hole Finder App"
frmGW.Caption = "Probable Gateway / Router IP"
frmPihole.Caption = "Pi-Hole Ad Blocker IP"
btnScan.Caption = "Find Gateway"
btnScan.Enabled = False
btnScanPi.Caption = "Find Pi-Hole"
btnScanPi.Enabled = False
txtStatus.Text = "INITIALIZING ENVIRONMENT..."      'clear the textbox
txtOutput.Text = ""                                 'clear the textbox

lblGW.Caption = "Unknown"
lblGW.Alignment = 2     '2 means centered
lblGW.FontUnderline = True
lblGW.FontStrikethru = True
lblGW.ForeColor = vbBlue

lblPiHole.Caption = "Unknown"
lblPiHole.Alignment = 2     '2 means centered
lblPiHole.FontUnderline = True
lblPiHole.FontStrikethru = True
lblPiHole.ForeColor = vbBlue

tmrCMD.Enabled = True
'txtOutput.Visible = False      'hides intermediate processing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
   If StdIO.Ready = True Then
    End
   Else
      bExitAfterCancel = True
      StdIO.Cancel
    End If
End Sub

Private Sub lblGW_Click()
ShellExecute 0, vbNullString, lblGW.Caption, vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub lblPiHole_Click()
ShellExecute 0, vbNullString, lblPiHole.Caption, vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub StdIO_CancelFail()  'Cancel failed to end program. No longer reading pipes.
    DoEvents
    If bExitAfterCancel Then End
End Sub

Private Sub StdIO_CancelSuccess()  'Cancel success! No longer reading pipes.
    DoEvents
    If bExitAfterCancel Then End
End Sub

Private Sub StdIO_Error(ByVal Number As Integer, ByVal Description As String)
    'Error #" & Number & ": " & Description
End Sub

Private Sub StdIO_GotData(ByVal Data As String)
    AddOutput Data
End Sub

Private Sub AddOutput(ByVal strData As String)
    txtOutput.Text = txtOutput.Text & strData
    txtOutput.SelStart = Len(txtOutput.Text)
End Sub

Private Sub tmrCMD_Timer()
txtStatus.Text = "READY"
txtOutput.Text = ""
btnScan.Enabled = True
btnScanPi.Enabled = True

If StdIO.Ready = True Then
   StdIO.CommandLine = "cmd"
   StdIO.ExecuteCommand             'Or simply StdIO.ExecuteCommand txtCommand.Text
End If

tmrCMD.Enabled = False

End Sub

Private Sub tmrGateway_Timer()
tmrGateway.Enabled = False
Call ProcessGateway
End Sub

Private Sub tmrPi_Timer()
tmrPi.Enabled = False
Call ProcessPi
End Sub
