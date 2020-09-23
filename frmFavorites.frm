VERSION 5.00
Begin VB.Form frmFavorites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Storm eX Favorites"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmFavorites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   1080
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAddSite 
      Caption         =   "Add Current Site"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtLink 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "http://"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox lstFavorites 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblNumber 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblNumberOfFavs 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Favorites:"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuMainRemove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The following (FileExists, Loadlistbox, and SaveListBox)
'were taken taken from the dos32.bas file.
Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Sub Loadlistbox(Directory As String, thelist As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        thelist.AddItem MyString$
    Wend
    Close #1
End Sub

Sub SaveListBox(Directory As String, thelist As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.List(SaveList&)
    Next SaveList&
    Close #1
End Sub

Private Sub cmdAdd_Click()
lstFavorites.AddItem txtLink.Text
End Sub

Private Sub cmdAddSite_Click()
lstFavorites.AddItem frmMain.cmbAddress.Text
End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrorHandler
lstFavorites.RemoveItem (lstFavorites.ListIndex)
Exit Sub

ErrorHandler:
 MsgBox "Please select a site to remove in the List and then select Remove", vbOKOnly, "Error In Removing Site"
End Sub

Private Sub Form_Load()
If FileExists("favorites.fav") = False Then
Dim sFile As String
Dim nFile As Integer
 nFile = FreeFile
 sFile = "favorites.fav"
  Open "favorites.fav" For Output As nFile
  Print #nFile, ""
  Close nFile
  lstFavorites.AddItem "http://stormgard5.tripod.com/"
        Exit Sub
End If
 Dim sTemp As String
 Open ("favorites.fav") For Input As #1
 While Not EOF(1)
  Line Input #1, sTemp
  If sTemp = "" Then GoTo End1
  lstFavorites.AddItem sTemp
 Wend
End1:
 Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveListBox "favorites.fav", lstFavorites
End Sub

Private Sub lstFavorites_DblClick()
frmMain.cmbAddress.Text = lstFavorites.Text
frmMain.Browser.Navigate frmMain.cmbAddress.Text
End Sub

Private Sub lstFavorites_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Me.PopupMenu mnuMain
End If
End Sub

Private Sub mnuMainRemove_Click()
cmdRemove_Click
End Sub

Private Sub Timer1_Timer()
lblNumber.Caption = lstFavorites.ListCount
End Sub
