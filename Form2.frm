VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "About TypePad 2001"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3960
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   0
      Picture         =   "Form2.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "RM01011982 - This is your software identification number. "
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "System User"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "This Program Is Licensed As A Single Program To:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   $"Form2.frx":0884
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   $"Form2.frx":0948
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "TypePad 2001 Build 0001RM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form2
End Sub
