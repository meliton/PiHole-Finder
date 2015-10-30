VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
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
   ScaleHeight     =   7365
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnScanPi 
      Caption         =   "btnScanPi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3960
      Width           =   2415
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
      Height          =   2775
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   1080
      Width           =   5895
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "btnExit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton btnScan 
      Caption         =   "btnScan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblPiholeMsg 
      Caption         =   "lblPiholeMsg"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label lblGWMsg 
      Caption         =   "lblGWMsg"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim WithEvents StdIO As cStdIO
Attribute StdIO.VB_VarHelpID = -1
Dim bExitAfterCancel As Boolean

Private Sub btnExit_Click()
Call cmdCancel  'kill the process gracefully before exiting
Unload Me
End Sub

Private Sub btnScan_Click()
Dim strNicConfig As String
txtOutput.Text = ""
strNicConfig = "wmic nicconfig WHERE IPEnabled=true GET DefaultIPGateway /format:csv"

If StdIO.Ready = True Then
   StdIO.CommandLine = strNicConfig    'runs the command to get the gateway
   StdIO.ExecuteCommand             'Or simply StdIO.ExecuteCommand txtCommand.Text
End If
Call cmdCancel      'kills the command prompt that is running
End Sub

Private Sub btnScanPi_Click()
Dim strPiHole As String
txtOutput.Text = ""
strPiHole = "net view"

If StdIO.Ready = True Then
   StdIO.CommandLine = strPiHole    'runs the command to get the gateway
   StdIO.ExecuteCommand             'Or simply StdIO.ExecuteCommand txtCommand.Text
End If
Call cmdCancel      'kills the command prompt that is running
End Sub

Private Sub Form_Load()
Set StdIO = New cStdIO
'txtOutput.Text = Environ("ComSpec")

'Set up app labels and buttons
Form1.Caption = "PiHole Finder App"
lblGWMsg.Caption = "Probable Gateway / Router IP(s)"
lblPiholeMsg.Caption = "Pi Hole Ad Blocker IP"
btnScan.Caption = "Find Gateway"
btnScanPi.Caption = "Find PiHole"
btnExit.Caption = "Exit"
txtOutput.Text = ""       'clear the textbox
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
