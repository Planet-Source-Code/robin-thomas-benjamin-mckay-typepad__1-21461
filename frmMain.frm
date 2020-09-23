VERSION 5.00
Object = "{249FCAA5-5488-4B89-B216-E05DC09DF237}#1.0#0"; "AXSSMTP.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Document Via Email"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   5880
      Width           =   735
   End
   Begin axsSMTP.axsSMTPSock axsSMTPSock1 
      Left            =   4200
      Top             =   5400
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server Information"
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   312
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Text            =   "smtp.somedomain.foo"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Server:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status Window:"
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   4455
      Begin VB.TextBox txtStatus 
         Height          =   1095
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "This is a test Email"
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtAttachments 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtRecipients 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtSender 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtBody 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress:"
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Message body:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Attachments - delimited by "";"":"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Recipients - delimited by "";"":"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sender Email Address:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sender Name:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub axsSMTPSock1_Connected()

    txtStatus.Text = txtStatus.Text & "Connected" & vbCrLf

    cmdConnect.Caption = "&Disconnect"
    cmdSend.Enabled = True
    txtServer.Enabled = False

End Sub
Private Sub axsSMTPSock1_Disconnected()
    
    ' Connection disconnected
    txtStatus.Text = txtStatus.Text & "Disconnected from server" & vbCrLf
        
    ' Re-enable all the buttons
    cmdConnect.Caption = "&Connect"
    cmdSend.Enabled = False
    txtServer.Enabled = True

End Sub
Private Sub axsSMTPSock1_Error(Number As Long, Description As String, ServerResponse As String)

    ' If it was an Invalid Response from the server then
    ' show what response the server sent.  Else
    ' show the error description.
    If Number = axsSMTPSockError_InvalidResponse Then
        txtStatus.Text = txtStatus.Text & "Error: " & ServerResponse & vbCrLf
    Else
        txtStatus.Text = txtStatus.Text & "Error: " & Description & vbCrLf
    End If

    ' Re-enable all the buttons just like in
    ' the "Disconnected" method because
    ' Once an error occurs, you will get
    ' disconnected from the server
    cmdConnect.Caption = "&Connect"
    cmdSend.Enabled = False
    txtServer.Enabled = True

End Sub
Private Sub axsSMTPSock1_InvalidRecipient(EmailAddress As String, ServerResponse As String, Reset As Boolean)

    Dim tempString As String

    tempString = "An error occured sending the following email address: " & EmailAddress & vbCrLf
    tempString = tempString & "The server response was: " & ServerResponse & vbCrLf
    tempString = tempString & "Would you like to continue sending?"
    
    ' An invalid recipient was specified.  Ask
    ' the user whether to continue sending the
    ' email or cancel it.  Once cancelled you
    ' should disconnect
    If MsgBox(tempString, vbYesNo) = vbYes Then
        Reset = False
    Else
        Reset = True
        DoEvents
        axsSMTPSock1.Disconnect
    End If

End Sub

Private Sub axsSMTPSock1_MailSendComplete()
    
    ' Once mail has been sent, disconnect
    ' from the server.
    txtStatus.Text = txtStatus.Text & "Send complete" & vbCrLf
    axsSMTPSock1.Disconnect

End Sub

Private Sub axsSMTPSock1_MessageProgress(MessageBytesSent As Long, TotalMessageSize As Long, CurrentAttachmentFilename As String)

    Dim tempString As String

    ' Display how much of the message has been sent
    tempString = MessageBytesSent & " of  " & TotalMessageSize & " bytes sent"
    If CurrentAttachmentFilename <> "" Then
        tempString = tempString & vbCrLf & "Current Attachment: " & CurrentAttachmentFilename
    End If

    lblProgress.Caption = tempString

End Sub

Private Sub cmdSend_Click()

    ' Make sure all the required information are filled in
    If txtEmail = "" Then
        MsgBox "No sender email address specified"
        Exit Sub
    End If
    If txtRecipients = "" Then
        MsgBox "No recipients specified"
        Exit Sub
    End If
    If txtSubject = "" Then
        MsgBox "No subject specified"
        Exit Sub
    End If
    If txtBody = "" Then
        MsgBox "No message specified"
        Exit Sub
    End If
    
    axsSMTPSock1.Attachments = txtAttachments
    axsSMTPSock1.MessageBody = txtBody
    axsSMTPSock1.MessageSubject = txtSubject
    axsSMTPSock1.Recipients = txtRecipients
    axsSMTPSock1.SenderEmailAddress = txtEmail
    axsSMTPSock1.SenderName = txtSender
    
    txtStatus.Text = txtStatus.Text & "Sending Email" & vbCrLf
    
    ' Send the email
    axsSMTPSock1.SendMail

End Sub

Private Sub cmdConnect_Click()
        
    If cmdConnect.Caption = "&Connect" Then
        ' Make sure the user specified a server
        ' or else a runtime error would occur
        If txtServer = "" Then
            MsgBox "No server specified"
            Exit Sub
        End If
        
        axsSMTPSock1.Server = txtServer
        txtStatus.Text = "Connecting to " & axsSMTPSock1.Server & vbCrLf
        
        'Connect to the server
        axsSMTPSock1.Connect
    Else
        axsSMTPSock1.Disconnect
    End If

End Sub

Private Sub Command1_Click()
Unload frmMain
End Sub

Private Sub Form_Load()
txtServer.Text = "smtp"
txtAttachments.Text = "c:\windows\temp.rtf"
End Sub
