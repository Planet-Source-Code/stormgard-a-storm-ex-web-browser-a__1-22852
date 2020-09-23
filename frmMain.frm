VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Storm eX Web Browser:"
   ClientHeight    =   5955
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FF8080&
      Caption         =   "Search"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00FF8080&
      Caption         =   "Home"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FF8080&
      Caption         =   "Refresh"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00FF8080&
      Caption         =   "Stop"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdForward 
      BackColor       =   &H00FF8080&
      Caption         =   ">>"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmbBack 
      BackColor       =   &H00FF8080&
      Caption         =   "<<"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox cmbAddress 
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      Text            =   "http://stormgard5.tripod.com/"
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00FF8080&
      Caption         =   "Go"
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   8745
      TabIndex        =   1
      Top             =   5700
      Width           =   8745
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   0
         Width           =   6375
         WordWrap        =   -1  'True
      End
   End
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2295
      ExtentX         =   4048
      ExtentY         =   3625
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileSavePageAs 
         Caption         =   "Save Page as..."
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuSettingsFont 
         Caption         =   "Font"
         Begin VB.Menu mnuSettingsFontMSSansSerif 
            Caption         =   "MS Sans Serif"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSettingsFontTimesNewRoman 
            Caption         =   "Times New Roman"
         End
         Begin VB.Menu mnuSettingsFontVerdana 
            Caption         =   "Verdana"
         End
      End
      Begin VB.Menu mnuSettingsBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingsSettings 
         Caption         =   "Settings..."
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "Favorites"
   End
   Begin VB.Menu mnuOther 
      Caption         =   "Other"
      Begin VB.Menu mnuOtherPageProperties 
         Caption         =   "Page Properties"
      End
      Begin VB.Menu mnuOtherBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtherViewSource 
         Caption         =   "View Source"
      End
      Begin VB.Menu mnuOtherBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtherTextPad 
         Caption         =   "Text Pad"
      End
      Begin VB.Menu mnuOtherAddressBook 
         Caption         =   "Address Book"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOtherBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtherMySite 
         Caption         =   "My Site"
      End
      Begin VB.Menu mnuOtherAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuOtherBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtherHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddressCombo(X As ComboBox)
On Error Resume Next
If Len(X) > 5 Then
 X.RemoveItem 5
End If
End Sub

Private Sub Browser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
cmbAddress.Text = URL
End Sub

Private Sub Browser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
ProgressBar1.Max = ProgressMax
ProgressBar1.Value = Progress
ProgressBar1.Refresh
End Sub

Private Sub Browser_StatusTextChange(ByVal Text As String)
lblStatus.Caption = Text
End Sub

Private Sub Browser_TitleChange(ByVal Text As String)
frmMain.Caption = "Storm eX Web Browser: " & Text & ""
Address = frmMain.Caption
End Sub

Private Sub cmbAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 cmdGo_Click
KeyAscii = 0
End If
End Sub

Private Sub cmbBack_Click()
On Error Resume Next
Browser.GoBack
Browser.SetFocus
End Sub

Private Sub cmdForward_Click()
On Error Resume Next
Browser.GoForward
Browser.SetFocus
End Sub

Private Sub cmdGo_Click()
On Error Resume Next
Browser.Navigate cmbAddress
cmbAddress.AddItem cmbAddress.Text
AddressCombo cmbAddress
Browser.SetFocus
End Sub

Private Sub cmdHome_Click()
On Error Resume Next
Browser.GoHome
Browser.SetFocus
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
Browser.Refresh
Browser.SetFocus
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
If frmSettings.txtSearchPage.Text = "" Then
 Browser.GoSearch
Else
 cmbAddress.Text = frmSettings.txtSearchPage.Text
 cmdGo_Click
End If
Browser.SetFocus
End Sub

Private Sub cmdStop_Click()
On Error Resume Next
Browser.Stop
Browser.SetFocus
End Sub

Private Sub Form_Load()
If frmSettings.txtStartupPage.Text = "" Then
 Browser.Navigate "about:blank"
Else
 cmbAddress.Text = frmSettings.txtStartupPage.Text
 cmdGo_Click
End If
frmMain.WindowState = 2
End Sub

Private Sub Form_Resize()
On Error Resume Next
Browser.Width = frmMain.Width - 130: Browser.Height = frmMain.Height - 1450
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmAbout
Unload frmFavorites
Unload frmSettings
Unload frmTextPad
Unload Me
End Sub

Private Sub mnuFavorites_Click()
frmFavorites.Visible = True
End Sub

Private Sub mnuFileExit_Click()
Unload frmAbout
Unload frmFavorites
Unload frmSettings
Unload frmTextPad
Unload Me
End Sub

Private Sub mnuFilePageSetup_Click()
On Error Resume Next
Browser.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuFilePrint_Click()
On Error Resume Next
Browser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuFilePrintPreview_Click()
On Error Resume Next
Browser.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuFileSavePageAs_Click()
On Error Resume Next
Browser.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuOtherAbout_Click()
frmAbout.Visible = True
End Sub

Private Sub mnuOtherHelp_Click()
cmbAddress.Text = "http://stormgard5.tripod.com/browser/index.html"
cmdGo_Click
End Sub

Private Sub mnuOtherMySite_Click()
cmbAddress.Text = "http://stormgard5.tripod.com/"
cmdGo_Click
End Sub

Private Sub mnuOtherPageProperties_Click()
On Error Resume Next
Browser.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuOtherTextPad_Click()
frmTextPad.Visible = True
End Sub

Private Sub mnuOtherViewSource_Click()
Dim addresstext As String
addresstext = cmbAddress.Text
cmbAddress.Text = "view-source:" & cmbAddress.Text
cmdGo_Click
cmbAddress.Text = addresstext
End Sub

Private Sub mnuSettingsFontMSSansSerif_Click()
mnuSettingsFontTimesNewRoman.Checked = False
mnuSettingsFontMSSansSerif.Checked = True
mnuSettingsFontVerdana.Checked = False
cmbAddress.Font = "MS Sans Serif"
lblStatus.Font = "MS Sans Serif"
End Sub

Private Sub mnuSettingsFontTimesNewRoman_Click()
mnuSettingsFontTimesNewRoman.Checked = True
mnuSettingsFontMSSansSerif.Checked = False
mnuSettingsFontVerdana.Checked = False
cmbAddress.Font = "Times New Roman"
lblStatus.Font = "Times New Roman"
End Sub

Private Sub mnuSettingsFontVerdana_Click()
mnuSettingsFontTimesNewRoman.Checked = False
mnuSettingsFontMSSansSerif.Checked = False
mnuSettingsFontVerdana.Checked = True
cmbAddress.Font = "Verdana"
lblStatus.Font = "Verdana"
End Sub

Private Sub mnuSettingsSettings_Click()
frmSettings.Visible = True
End Sub
