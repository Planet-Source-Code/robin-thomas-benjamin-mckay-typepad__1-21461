VERSION 5.00
Begin VB.Form frmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Margins setup"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "frmPageSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2730
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   915
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   270
      TabIndex        =   2
      Top             =   90
      Width           =   3390
      Begin VB.ComboBox cboBottomMargin 
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1755
         Width           =   1140
      End
      Begin VB.ComboBox cboTopMargin 
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1305
         Width           =   1140
      End
      Begin VB.ComboBox cboRightMargin 
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   855
         Width           =   1140
      End
      Begin VB.ComboBox cboLeftMargin 
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   405
         Width           =   1140
      End
      Begin VB.Label lblBottomMargin 
         Caption         =   "Bottom   (in)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label lblTopMargin 
         Caption         =   "Top   (in)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1395
         Width           =   1350
      End
      Begin VB.Label lblRightMargin 
         Caption         =   "Right   (in)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label lblLeftMargin 
         Caption         =   "Left    (in)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   450
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1500
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Width           =   915
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PageSetup.frm
'
' By Herman Liu

Option Explicit



Private Sub Form_Load()
    gprint = False
    Dim i, m
    m = 4
    For i = 0 To m Step 0.05
        frmPageSetup.cboLeftMargin.AddItem FormatNumber(i, 2)
    Next
    For i = 0 To m Step 0.05
        frmPageSetup.cboRightMargin.AddItem FormatNumber(i, 2)
    Next
    For i = 0 To m Step 0.05
        frmPageSetup.cboTopMargin.AddItem FormatNumber(i, 2)
    Next
    For i = 0 To m Step 0.05
        frmPageSetup.cboBottomMargin.AddItem FormatNumber(i, 2)
    Next
    frmPageSetup.cboLeftMargin.Text = cboLeftMargin.List(gLeftMargin / 0.05)
    frmPageSetup.cboRightMargin.Text = cboRightMargin.List(gRightMargin / 0.05)
    frmPageSetup.cboTopMargin.Text = cboTopMargin.List(gTopMargin / 0.05)
    frmPageSetup.cboBottomMargin.Text = cboBottomMargin.List(gBottomMargin / 0.05)
End Sub


Private Sub cmdOK_Click()
    gLeftMargin = cboLeftMargin.Text
    gRightMargin = cboRightMargin.Text
    gTopMargin = cboTopMargin.Text
    gBottomMargin = cboBottomMargin.Text
    Unload Me
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdPrint_Click()
    gprint = True
    Unload Me
End Sub
