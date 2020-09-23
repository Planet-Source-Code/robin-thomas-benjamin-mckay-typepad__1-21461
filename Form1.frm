VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Date/Time"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time/Date"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmDocMaster.Text2.Text = List1.Text
Unload Form1
Screen.ActiveForm.ActiveControl.SelText = frmDocMaster.Text2.Text
frmDocMaster.Text2.Text = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
List1.AddItem Format(Now, "m/d/yy")
List1.AddItem Format(Now, "dddd, mmmm dd")
List1.AddItem Format(Now, "d-mmm")
List1.AddItem Format(Now, "mmmm-yy")
List1.AddItem Format(Now, "hh:mm AM/PM")
List1.AddItem Format(Now, "h:mm:ss a/p")
List1.AddItem Format(Now, "d-mmmm h:mm")
List1.AddItem Format(Now, "ddddd ttttt")
End Sub
