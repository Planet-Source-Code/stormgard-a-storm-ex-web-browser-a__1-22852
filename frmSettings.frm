VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Storm eX Settings"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveSettings 
      Caption         =   "Save Settings"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdBlank 
      Caption         =   "Use Blank"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtStartupPage 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "about:blank"
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtSearchPage 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "http://www.google.com/"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblStartupPage 
      Alignment       =   1  'Right Justify
      Caption         =   "Start-UpPage:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblSearchPage 
      Alignment       =   1  'Right Justify
      Caption         =   "Search Page:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBlank_Click()
txtStartupPage.Text = "about:blank"
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSaveSettings_Click()
Call SaveSetting("Storm eX Web Browser", "Settings", "Start-Up Page", txtStartupPage.Text)
Call SaveSetting("Storm eX Web Browser", "Settings", "Search Page", txtSearchPage.Text)
Unload Me
End Sub

Private Sub Form_Load()
Line1.X1 = frmSettings.Left
Line1.X2 = frmSettings.Width
txtStartupPage.Text = GetSetting("Storm eX Web Browser", "Settings", "Start-Up Page", "")
txtSearchPage.Text = GetSetting("Storm eX Web Browser", "Settings", "Search Page", "")
End Sub
