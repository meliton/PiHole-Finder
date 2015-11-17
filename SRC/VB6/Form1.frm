VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmPihole 
      Caption         =   "frmPihole"
      Height          =   2055
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton btnScanPi 
         Caption         =   "btnScanPi"
         Height          =   495
         Left            =   1080
         TabIndex        =   6
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblPiHole 
         Caption         =   "lblPiHole"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   600
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
      Height          =   1245
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   2280
      Width           =   9975
   End
   Begin VB.Frame frmGW 
      Caption         =   "frmGW"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton btnScan 
         Caption         =   "btnScan"
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblGW 
         Caption         =   "lblGW"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3735
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
Dim strGateway As String
txtOutput.Text = ""
'strGateway = "ipconfig | find " & Chr(34) & "Default" & Chr(34)
strGateway = "findGW.bat"

If StdIO.Ready = True Then
   StdIO.CommandLine = strGateway   'runs the command to get the gateway
   StdIO.ExecuteCommand             'Or simply StdIO.ExecuteCommand txtCommand.Text
End If

txtOutput.Text = Replace$(txtOutput.Text, "Default Gateway", "")   'removes default gateway text
txtOutput.Text = Replace$(txtOutput.Text, " . ", "")    'replace period/spaces with null
txtOutput.Text = Replace$(txtOutput.Text, ": ", "")     'replace colon with null
txtOutput.Text = Replace$(txtOutput.Text, "..", "")     'removes extra periods with null

If txtOutput.Text = "" Then         'check here if string is null. If null gateway was not found so exit sub
lblGW.Caption = "NOT FOUND"
txtOutput.Text = "STATUS... Gateway was not found. Your network is not up."
Exit Sub
Else
End If

txtOutput.Text = Trim$(txtOutput.Text)      'removes leading and trailing spaces
txtOutput.Text = LineTrim(txtOutput.Text)   'removes blank lines

lblGW.Caption = "http://" & txtOutput.Text
lblGW.FontStrikethru = False

Call cmdCancel      'kills the command prompt that is running

End Sub

Private Sub btnScanPi_Click()
Dim strPiHole As String
txtOutput.Text = ""
strPiHole = "findPi.bat "   'scan for Raspberry MACs like B8:27:EB

If StdIO.Ready = True Then
   StdIO.CommandLine = strPiHole    'runs the command to get the gateway
   StdIO.ExecuteCommand             'Or simply StdIO.ExecuteCommand txtCommand.Text
End If
Call cmdCancel      'kills the command prompt that is running

txtOutput.Text = Left$(txtOutput.Text, 21)  'get the first 21 characters of the textbox
txtOutput.Text = Trim$(txtOutput.Text)      'removes leading and trailing spaces

If txtOutput.Text = "" Then         'check here if string is null. If null pi was not found so exit sub
lblPiHole.Caption = "NOT FOUND"
txtOutput.Text = "STATUS... Pi is not found. Try a ping sweep to wake up Pi"
Exit Sub
Else
End If

txtOutput.Text = LineTrim(txtOutput.Text)   'removes blank lines
lblPiHole.Caption = "http://" & txtOutput.Text & "/admin"
lblPiHole.FontStrikethru = False

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
Form1.Caption = "PiHole Finder App"
frmGW.Caption = "Probable Gateway / Router IP"
frmPihole.Caption = "Pi Hole Ad Blocker IP"
btnScan.Caption = "Find Gateway"
btnScanPi.Caption = "Find PiHole"
txtOutput.Text = "STATUS..."       'clear the textbox

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

End Sub

Private Sub cmdCancel()
If StdIO.Ready = False Then
    StdIO.Cancel
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    If StdIO.Ready = True Then
        End
    Else
        bExitAfterCancel = True
        cmdCancel
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
