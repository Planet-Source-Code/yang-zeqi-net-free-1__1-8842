VERSION 5.00
Begin VB.Form frmIO 
   Caption         =   "Internet Options"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "frmIO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtHP 
      Height          =   285
      Left            =   480
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Text            =   "www.planet-source-code.com"
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblHP 
      Caption         =   "Home Page:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    HP = txtHP.Text
    Unload Me
End Sub
