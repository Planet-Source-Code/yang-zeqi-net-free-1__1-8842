VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Net Free 1"
   ClientHeight    =   4830
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   6435
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   8280
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox cmbAddress 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   6495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   7335
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   9615
      ExtentX         =   16960
      ExtentY         =   12938
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
   Begin VB.ListBox lstSites 
      Height          =   6885
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Favourites"
      Top             =   1200
      Width           =   2175
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   240
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0784
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":102A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":130C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrToolbar 
      Height          =   510
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   900
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFavourites 
      Caption         =   "Favourites:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsAf 
         Caption         =   "Add to &Favourites"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsInternetOptions 
         Caption         =   "&Internet Options"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HP As String

Private Sub cmdGo_Click()
    WebBrowser.Navigate cmbAddress.Text
    cmbAddress.AddItem cmbAddress.Text
End Sub

Private Sub Form_Load()
    HP = frmIO.txtHP.Text
    cmbAddress.Text = HP
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dialog.Show
End Sub

Private Sub lstSites_Click()
    'jump to the selected site
    WebBrowser.Navigate _
        Trim$(lstSites.Text)
End Sub

Private Sub mnuFileQuit_Click()
    Dialog.Show
End Sub

Private Sub mnuOptionsAf_Click()
    lstSites.AddItem cmbAddress.Text
End Sub

Private Sub mnuOptionsInternetOptions_Click()
    frmIO.Show
End Sub

Private Sub tbrToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    'skip error checking
    On Error Resume Next
    
    'what button was pressed?
    Select Case UCase$(Trim(Button.Key))
        Case Is = "BACK"
            WebBrowser.GoBack
        Case Is = "Forward"
            WebBrowser.GoForward
        Case Is = "refresh"
            WebBrowser.Refresh
        Case Is = "HOME"
            WebBrowser.Navigate HP
            cmbAddress.Text = HP
        End Select
End Sub
