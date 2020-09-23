VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{97F4CED3-9103-11CE-8385-524153480001}#2.0#0"; "VSPELL32.OCX"
Begin VB.Form frmDocMaster 
   Caption         =   "Document "
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7530
   Icon            =   "frmDocMaster.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   7530
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmDocMaster.frx":0442
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   13150
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDocMaster.frx":04F0
   End
   Begin VspelocxLib.VSSpell VSSpell1 
      Left            =   3600
      OleObjectBlob   =   "frmDocMaster.frx":059E
      Top             =   1680
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   3480
      Top             =   1680
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu border3 
         Caption         =   "-"
      End
      Begin VB.Menu Open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu border1 
         Caption         =   "-"
      End
      Begin VB.Menu PAGESETUP 
         Caption         =   "&Page Setup"
      End
      Begin VB.Menu Preview 
         Caption         =   "&Print Preview "
      End
      Begin VB.Menu Pirnt 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu border2 
         Caption         =   "-"
      End
      Begin VB.Menu Send 
         Caption         =   "Sen&d..."
      End
      Begin VB.Menu newborder 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu Del 
         Caption         =   "&Delete"
      End
      Begin VB.Menu border4 
         Caption         =   "-"
      End
      Begin VB.Menu SELECTALL 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu TIMEDATE 
         Caption         =   "&Time/Date"
      End
   End
   Begin VB.Menu Insert 
      Caption         =   "&Insert"
      Begin VB.Menu Piccy 
         Caption         =   "&Picture"
         Shortcut        =   {F9}
      End
      Begin VB.Menu Filey 
         Caption         =   "&File"
         Shortcut        =   {F11}
      End
      Begin VB.Menu TFB 
         Caption         =   "-"
      End
      Begin VB.Menu TFILE 
         Caption         =   "&Text File"
         Shortcut        =   {F12}
      End
      Begin VB.Menu brandnewbo 
         Caption         =   "-"
      End
      Begin VB.Menu Font 
         Caption         =   "&Font"
      End
      Begin VB.Menu brandnewb 
         Caption         =   "-"
      End
      Begin VB.Menu Bullet 
         Caption         =   "&Bullet"
      End
      Begin VB.Menu bulletborder 
         Caption         =   "-"
      End
      Begin VB.Menu Colour 
         Caption         =   "&Colour"
      End
   End
   Begin VB.Menu Agent 
      Caption         =   "&Agent"
      Begin VB.Menu DICTATEDOC 
         Caption         =   "&Dictate Document..."
      End
   End
   Begin VB.Menu Format 
      Caption         =   "&Format"
      Begin VB.Menu Paragraph 
         Caption         =   "&Paragraph"
         Begin VB.Menu Left 
            Caption         =   "&Left "
         End
         Begin VB.Menu Right 
            Caption         =   "&Right"
         End
         Begin VB.Menu Center 
            Caption         =   "&Center"
         End
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
      Begin VB.Menu SpellCheck 
         Caption         =   "&Spell Check All Text"
         Shortcut        =   {F5}
      End
      Begin VB.Menu SCSW 
         Caption         =   "&Spell Check Single Word"
      End
      Begin VB.Menu spswb 
         Caption         =   "-"
      End
      Begin VB.Menu newemail 
         Caption         =   "&New Email"
      End
      Begin VB.Menu newemailborder 
         Caption         =   "-"
      End
      Begin VB.Menu View 
         Caption         =   "&View"
         Begin VB.Menu InMicrosoftWord 
            Caption         =   "&In Microsoft Word..."
         End
      End
   End
   Begin VB.Menu Window 
      Caption         =   "&Window"
      Begin VB.Menu Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu THORIZONTAL 
         Caption         =   "&Tile Horizontal"
      End
      Begin VB.Menu TVERTICAL 
         Caption         =   "&Tile Vertical"
      End
      Begin VB.Menu ARRANGEICONS 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu HELP 
      Caption         =   "&Help"
      Begin VB.Menu EMAILTHEAUTHOR 
         Caption         =   "&Email The Author"
      End
      Begin VB.Menu bnborder 
         Caption         =   "-"
      End
      Begin VB.Menu HowDoI 
         Caption         =   "&How Do I"
         Begin VB.Menu PRINTPREVIEWH 
            Caption         =   "&Print Preview"
         End
         Begin VB.Menu PMYDOC 
            Caption         =   "&Print My Document"
         End
         Begin VB.Menu SMDOC 
            Caption         =   "&Send My Document"
         End
         Begin VB.Menu SpellCheck2 
            Caption         =   "&Spell Check"
         End
         Begin VB.Menu spbor 
            Caption         =   "-"
         End
         Begin VB.Menu FOTOPIC 
            Caption         =   "&Find Other Topic..."
         End
      End
      Begin VB.Menu HELPTOPICS 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu ABO 
         Caption         =   "-"
      End
      Begin VB.Menu Credits 
         Caption         =   "&Credits"
      End
      Begin VB.Menu aepb 
         Caption         =   "-"
      End
      Begin VB.Menu AEP2001 
         Caption         =   "&About TypePad 2001"
      End
   End
End
Attribute VB_Name = "frmDocMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub AEP2001_Click()
Form2.Show
End Sub

Private Sub ARRANGEICONS_Click()
frmFrame.Arrange vbArrangeIcons
End Sub

Private Sub Bullet_Click()
If Screen.ActiveForm.ActiveControl.SelBullet = False Then
    Screen.ActiveForm.ActiveControl.SelBullet = True
    Bullet.Checked = True
Else
If Screen.ActiveForm.ActiveControl.SelBullet = True Then
    Screen.ActiveForm.ActiveControl.SelBullet = False
    Bullet.Checked = False
End If
End If
End Sub

Private Sub Cascade_Click()
frmFrame.Arrange vbCascade
End Sub

Private Sub Center_Click()
Screen.ActiveForm.ActiveControl.SelAlignment = rtfCenter
End Sub

Private Sub Colour_Click()
With frmFrame.CommonDialog1
    .CancelError = True
    .Flags = cdlCCFullOpen
    .ShowColor
End With
Screen.ActiveForm.ActiveControl.SelColor = frmFrame.CommonDialog1.Color
End Sub

Private Sub Copy_Click()
Clipboard.SetText Screen.ActiveForm.ActiveControl.SelText
End Sub

Private Sub Credits_Click()
frmCredits.Show
End Sub

Private Sub Cut_Click()
Clipboard.SetText Screen.ActiveForm.ActiveControl.SelText
Screen.ActiveForm.ActiveControl.SelText = ""
End Sub

Private Sub Del_Click()
If Screen.ActiveForm.ActiveControl.SelText = "" Then
    MsgBox "Nothing to delete", vbInformation, "Error:"
Else
    Screen.ActiveForm.ActiveControl.SelText = ""
End If
End Sub

Private Sub DICTATEDOC_Click()
On Error Resume Next
Set hanz = Agent1.Characters.Character("Hanz")
hanz.Show
hanz.Speak Screen.ActiveForm.ActiveControl.Text
hanz.Hide
End Sub



Private Sub EMAILTHEAUTHOR_Click()
Shell ("start mailto:ian@imckay.fsnet.co.uk"), vbHide
End Sub

Private Sub Exit_Click()
Unload frmFrame
End Sub

Private Sub Filey_Click()
On Error Resume Next
With frmFrame.CommonDialog1
    .CancelError = True
    .Filter = "All Files|*.*"
    .ShowOpen
End With
Screen.ActiveForm.ActiveControl.OLEObjects.Add , , frmFrame.CommonDialog1.FileName
End Sub

Private Sub Font_Click()
On Error Resume Next
With frmFrame.CommonDialog1
    .CancelError = True
    .Flags = cdlCFBoth + cdlCFEffects
    .ShowFont
End With
Screen.ActiveForm.ActiveControl.Font.Name = frmFrame.CommonDialog1.FontName
Screen.ActiveForm.ActiveControl.Font.Size = frmFrame.CommonDialog1.FontSize
Screen.ActiveForm.ActiveControl.SelBold = frmFrame.CommonDialog1.FontBold
Screen.ActiveForm.ActiveControl.SelItalic = frmFrame.CommonDialog1.FontItalic
Screen.ActiveForm.ActiveControl.SelUnderline = frmFrame.CommonDialog1.FontUnderline
Screen.ActiveForm.ActiveControl.SelStrikeThru = frmFrame.CommonDialog1.FontStrikethru
Screen.ActiveForm.ActiveControl.SelColor = frmFrame.CommonDialog1.Color
End Sub

Private Sub Form_Load()
Agent1.Characters.Load "Hanz", "Hanz.acs"
Set hanz = Agent1.Characters.Character("Hanz")
hanz.Hide
End Sub
Private Sub form_unload(cancel As Integer)
On Error Resume Next
Dim rtb1 As Integer
    rtb1 = MsgBox("Would you like to save the changes to this document?", vbYesNo + vbQuestion, "Save Changes?")
        Select Case rtb1
            Case vbYes
                With frmFrame.CommonDialog1
                    .CancelError = True
                    .Filter = "RTF Files|*.rtf|TXT Files|*.txt|WRI Files|*.wri"
                    .ShowSave
                End With
                Screen.ActiveForm.ActiveControl.SaveFile (frmFrame.CommonDialog1.FileName)
            Case vbNo
                Exit Sub
        End Select
End Sub

Private Sub FOTOPIC_Click()
Form3.Show
End Sub

Private Sub HELPTOPICS_Click()
Form3.Show
End Sub

Private Sub InMicrosoftWord_Click()
On Error Resume Next
Set word = CreateObject("word.basic")

word.appshow
word.filenew
word.Insert Screen.ActiveForm.ActiveControl.Text
End Sub

Private Sub Left_Click()
Screen.ActiveForm.ActiveControl.SelAlignment = rtfLeft
End Sub

Private Sub New_Click()
On Error Resume Next
Dim frmDocMaster As New frmDocMaster
    frmDocMaster.Show
End Sub





Private Sub NEWEMAIL_Click()
Shell ("START MAILTO:"), vbHide
End Sub

Private Sub Open_Click()
' Opens up a file for principal viewing
On Error Resume Next
With frmFrame.CommonDialog1
    .CancelError = True
    .Filter = "RTF Files, TXT Files, WRI Files|*.wri; *.txt; *.rtf;"
    .ShowOpen
End With
Screen.ActiveForm.Caption = "Document -" + frmFrame.CommonDialog1.FileName
Screen.ActiveForm.ActiveControl.LoadFile (frmFrame.CommonDialog1.FileName)
End Sub

Private Sub PAGESETUP_Click()
' Here is the Page Setup Option
    frmPageSetup.Show vbModal
    If gprint = True Then
         frmDocPreview.DocPrintProc
    End If
End Sub



Private Sub Paste_Click()
Screen.ActiveForm.ActiveControl.SelText = Clipboard.GetText
End Sub

Private Sub Piccy_Click()
On Error Resume Next
With frmFrame.CommonDialog1
    .CancelError = True
    .Filter = "BMP Files, JPG Files, JPEG Files |*.bmp; *.jpg; *.jpeg;|All Files|*.*"
    .ShowOpen
End With
Screen.ActiveForm.ActiveControl.OLEObjects.Add , , frmFrame.CommonDialog1.FileName
End Sub

Private Sub Pirnt_Click()
frmDocPreview.DocPrintProc
End Sub

Private Sub PMYDOC_Click()
Set hanz = Agent1.Characters.Character("Hanz")
hanz.Show
hanz.Speak "To Print a document, please go to the file menu and then click Print. Press OK to start printing or cancel to abort."
hanz.Hide
End Sub

Private Sub Preview_Click()
    frmDocPreview.Show vbModal
    If gprint = True Then
         frmDocPreview.DocPrintProc
    End If
End Sub

Private Sub PRINTPREVIEWH_Click()
Set hanz = Agent1.Characters.Character("Hanz")
hanz.Show
hanz.Speak "Please go to the file menu and then click Print Preview. The document shall then be displayed."
hanz.Hide
End Sub

Private Sub Right_Click()
Screen.ActiveForm.ActiveControl.SelAlignment = rtfRight
End Sub

Private Sub Save_Click()
On Error Resume Next
With frmFrame.CommonDialog1
    .CancelError = True
    .Filter = "RTF Files|*.rtf|TXT Files|*.txt|WRI Files|*.wri"
    .ShowSave
End With
Screen.ActiveForm.ActiveControl.SaveFile (frmFrame.CommonDialog1.FileName)
End Sub

Private Sub SCSW_Click()
' Spell checks a single word in the document
frmDocMaster.VSSpell1.CheckText = Screen.ActiveForm.ActiveControl.SelText
If Screen.ActiveForm.ActiveControl.SelText = "" Then
    MsgBox "No text to check", vbInformation, "Error"
    Exit Sub
End If
If frmDocMaster.VSSpell1.ReplaceOccurred Then
    Screen.ActiveForm.ActiveControl.SelText = frmDocMaster.VSSpell1.Text
Else
    MsgBox "The Spell Check Is Complete", vbInformation, "The Spell Check Is Complete!"
End If
End Sub

Private Sub SELECTALL_Click()
Screen.ActiveForm.ActiveControl.SelStart = 0
Screen.ActiveForm.ActiveControl.SelLength = Len(Screen.ActiveControl)
End Sub

Private Sub Send_Click()
frmMain.Show
End Sub

Private Sub SMDOC_Click()
Set hanz = Agent1.Characters.Character("Hanz")
hanz.Show
hanz.Speak "Go to the File Menu and then click Send. This will display a send screen. Please fill in all your details and then click Send. Ensure that you are connected to the net before you do this. You must also know what your SMTP address is as well."
hanz.Hide
End Sub

Private Sub SPELLCHECK_Click()
frmDocMaster.VSSpell1.CheckText = Screen.ActiveForm.ActiveControl.Text
If frmDocMaster.VSSpell1.ReplaceOccurred Then
    Screen.ActiveForm.ActiveControl.Text = frmDocMaster.VSSpell1.Text
    MsgBox "The Spell Check Is Complete.", vbInformation, "The Spell Check Is Complete!"
Else
    Exit Sub
End If
End Sub

Private Sub SpellCheck2_Click()
Set hanz = Agent1.Characters.Character("Hanz")
hanz.Show
hanz.Speak "Go to Tools and then click Spell Checker. This will find any misspelled words and give you the chance to replace them. All words should be replaced and then you can return to the document."
hanz.Hide
End Sub

Private Sub Text1_Change()
Screen.ActiveForm.ActiveControl.SaveFile ("c:\windows\temp.rtf")
End Sub

Private Sub TFILE_Click()
On Error Resume Next
With frmFrame.CommonDialog1
    .CancelError = True
    .DialogTitle = "Insert Text File..."
    .Filter = "Text Files|*.txt"
    .ShowOpen
End With
frmDocMaster.RichTextBox1.Text = ""
frmDocMaster.RichTextBox1.LoadFile (frmFrame.CommonDialog1.FileName)
Screen.ActiveForm.ActiveControl.SelText = frmDocMaster.RichTextBox1.Text
End Sub

Private Sub THORIZONTAL_Click()
frmFrame.Arrange vbTileHorizontal
End Sub

Private Sub TIMEDATE_Click()
Form1.Show
End Sub

Private Sub TVERTICAL_Click()
frmFrame.Arrange vbTileVertical
End Sub
