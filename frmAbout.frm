VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Storm eX About"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Caption         =   "About:"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version: "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
 Unload Me
End Sub

Private Sub Form_Load()
lblTitle.Caption = "App Title: Storm eX Web Browser"
lblVersion.Caption = "Version: 2.2"
lblAbout.Caption = "About App: " & vbCrLf & "     I created the Storm eX Web Browser with visual basic 6 pro. It uses microsoft's internet controls and microsoft common controls." & vbCrLf & "http://stormgard5.tripod.com/" & vbCrLf & "stormgard5@yahoo.com"
End Sub
