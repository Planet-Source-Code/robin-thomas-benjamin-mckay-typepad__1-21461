VERSION 5.00
Begin VB.Form frmCredits 
   Caption         =   "Credits For TypePad 2001"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4440
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   $"frmCredits.frx":0442
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   3600
      Width           =   6255
   End
   Begin VB.Label Label5 
      Caption         =   $"frmCredits.frx":04CF
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   6255
   End
   Begin VB.Label Label4 
      Caption         =   $"frmCredits.frx":056B
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   6255
   End
   Begin VB.Label Label3 
      Caption         =   $"frmCredits.frx":0603
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   $"frmCredits.frx":06AC
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCredits.frx":0756
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmCredits
End Sub
