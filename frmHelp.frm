VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   Caption         =   "Help Topics"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4380
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6588
      _Version        =   393217
      BackColor       =   65535
      Enabled         =   0   'False
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmHelp.frx":0442
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Unload frmhelp
End Sub

Private Sub Form_Load()
frmhelp.Caption = "Help Topics -" + Form3.List1.Text
End Sub
