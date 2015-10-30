VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
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
   ScaleHeight     =   6105
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
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
      Left            =   3840
      TabIndex        =   1
      Top             =   5520
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
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   1335
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
      Top             =   3000
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
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnScan_Click()
Dim nicConfig As String
MsgBox "wmic NICCONFIG WHERE IPEnabled=true GET DefaultIPGateway /format:csv"
End Sub

Private Sub Form_Load()

'Set up app labels and buttons
Form1.Caption = "PiHole Finder App"
lblGWMsg.Caption = "Probable Gateway / Router IP(s)"
lblPiholeMsg.Caption = "Pi Hole Ad Blocker IP"
btnScan.Caption = "Scan"
btnExit.Caption = "Exit"
End Sub
