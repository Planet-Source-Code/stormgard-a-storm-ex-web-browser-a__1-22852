VERSION 5.00
Begin VB.Form frmTextPad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Storm eX Text Pad"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmTextPad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBox 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmTextPad.frx":030A
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmTextPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
txtBox.Text = GetSetting("Storm eX Web Browser", "Text Pad", "Notes", "")
If txtBox.Text = "" Then
 txtBox.Text = "You can type little notes to yourself in here. When this form is closed it will automatically save your notes to the system registry. Read the readme file to learn about clearing the registry"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SaveSetting("Storm eX Web Browser", "Text Pad", "Notes", txtBox.Text)
End Sub
