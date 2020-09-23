VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form3 
   Caption         =   "Help Topics"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7290
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Email Me If You Want To See Other Topics"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   6480
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "You only need to click a help topic once to bring it up"
      Top             =   480
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6255
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Help Topics"
      TabPicture(0)   =   "Form3.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form3
End Sub

Private Sub Command2_Click()
Shell ("start mailto:ian@imckay.fsnet.co.uk"), vbHide
End Sub

Private Sub Form_Load()
List1.AddItem "Cancelling Dialogs"
List1.AddItem "How To Create A New File"
List1.AddItem "How To Open A File"
List1.AddItem "How To Save A File"
List1.AddItem "How To Execute Page Setup"
List1.AddItem "How To Execute Print Preview"
List1.AddItem "How To Execute Print"
List1.AddItem "How To Send A Document Via Email"
List1.AddItem "How To Terminate The Application"
List1.AddItem "Cut To The Clipboard"
List1.AddItem "Copy To The Clipboard"
List1.AddItem "Paste From The Clipboard"
List1.AddItem "Delete Text"
List1.AddItem "Select All Text In Your Document"
List1.AddItem "Input Time And Date Into Your Document"
List1.AddItem "Working With Pictures"
List1.AddItem "Working With Files"
List1.AddItem "Working With Text Files"
List1.AddItem "Working With Fonts"
List1.AddItem "Working With Colours"
List1.AddItem "Working With Your Friend Hanz"
List1.AddItem "Paragraph Alignment"
List1.AddItem "Spell Check"
List1.AddItem "Initiate A New Email"
List1.AddItem "View In Microsoft Word"
List1.AddItem "Changing The View"
List1.AddItem "Working With Help"
End Sub

Private Sub List1_Click()
If List1.Text = "Cancelling Dialogs" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "When a dialog such as an Open dialog or Save dialog is displayed, press Cancel at any time to cancel the dialog and return to your document."
ElseIf List1.Text = "How To Create A New File" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "To Create a new file, go to the File Menu and Click New. This will create a new document for you."
ElseIf List1.Text = "How To Open A File" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "To Open A File, go to the file menu and click open. The Open Dialog will then appear. Browse through the directories by double clicking them until you find the file you are looking for."
ElseIf List1.Text = "How To Save A File" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "To Save A File, go to the file menu and click save. Change the filter at the bottom of the dialog box to save your document as either a Text file(txt), a rich text file(rtf) or a wordpad file(wri). Once you have made your choice, click save after you have typed a filename for your document."
ElseIf List1.Text = "How To Execute Page Setup" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "Go to the file menu and click page setup. You will then be presented with a Page Setup dialog box. You can change how you want your document printed by altering the margin properties. Click OK once you are done, or Print to print the document or cancel to return to your document."
ElseIf List1.Text = "How To Execute Print Preview" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "Go to the file menu and click Print Preview to see what your document will be like when it is printed. Change the zoom factor to 100 to make the document more clearer."
ElseIf List1.Text = "How To Execute Print" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "Go To the file menu and click Print. Press OK if you accept the settings or cancel if you want to return to your document straight away."
ElseIf List1.Text = "How To Send A Document Via Email" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "Go to the file menu and click Send. A dialog will appear and ask you to enter details of your e-mail account. The document is already attached. You can attach more documents, separating them with a ';' . Click Connect once you have filled in all your details, regardless of whether or not you are Online. Then Click Send. Your document will then be sent through the e-mail system."
ElseIf List1.Text = "How To Terminate The Application" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "Go to the file menu and click exit. Remember to SAVE your work before you exit the program. However, you will be prompted to save your work before you quit. Make sure you remember to do so if you need it to be saved."
ElseIf List1.Text = "Cut To The Clipboard" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "Go to the Edit menu and click Cut after selecting a word or sentence. To select a word or sentence, you need to highlight it by dragging your cursor over the word. Press Cut to complete the operation. Alternatively, you can press CTRL then X to cut your selected text to the clipboard."
ElseIf List1.Text = "Copy To The Clipboard" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.SelText = "Go to the Edit menu and click Copy. Repeat the Copy process to copy text to the clipboard. A simpler way is to press CTRL and then V."
ElseIf List1.Text = "Paste From The Clipboard" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Press CTRL and then V to paste from the clipboard."
ElseIf List1.Text = "Delete Text" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Select the string you wish to delete then click Delete in the Edit Menu. The selected string will then disappear completely. It is not retrievable so once it is deleted, it is gone forever. Be careful when executing this command. To unselect a string, just click anywhere else in the document, preferably as far away from the selected string as possible."
ElseIf List1.Text = "Select All Text In Your Document" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Press CTRL and then A to select all text in your document."
ElseIf List1.Text = "Input Time And Date Into Your Document" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the edit menu and click Time/Date. This will bring up the Time/Date dialog. Select a date and then press OK to insert the date into your document."
ElseIf List1.Text = "Working With Pictures" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Insert menu and then click Picture. A dialog box will appear. Select a picture to insert into your document. To actually insert a picture, press OK. WARNING: Large pictures may cause the program to halt."
ElseIf List1.Text = "Working With Files" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Insert menu and click file. Select a file to insert and then press OK. The file will be inserted into your document. This is known as embedding. You can do this as many times as you want to. There is no limit. Press OK to actually insert a file into your documents."
ElseIf List1.Text = "Working With Text Files" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Insert menu and click Text File. A dialog box will be displayed, allowing access only to text files. Select a text file you wish to insert and then press OK. The text file will be inserted at the place where the insertion point is."
ElseIf List1.Text = "Working With Fonts" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Insert menu and select font. Choose the font attributes that best suit you and then click OK."
ElseIf List1.Text = "Working With Colours" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Insert menu and select Colour. You can change the brightness of the font you select from the colour dialog box as well. Press OK when done."
ElseIf List1.Text = "Working With Your Friend Hanz" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Agent menu and then click Dictate Document. Hanz will then read out the entire contents of your document to you. However, some words might not be pronounced properly, depending on the complexity of the word. Please bear this in mind when doing this."
ElseIf List1.Text = "Paragraph Alignment" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Format menu, highlight paragraph and then wait for the next screen to appear. You can change the alignment by clicking left right or center. Clicking Left aligns text to the left, Clicking Center aligns text in the middle and clicking right aligns text to the right."
ElseIf List1.Text = "Spell Check" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Tools menu and click Spell Check All Text. This will check text for misspelled words and give you the option to replace them. However, you can check a single word by clicking Spell Check Single Word in the same menu. Once the operation is complete, a dialog will appear saying the spell check is complete. On some systems, the VCI VisualSpeller control might not be evident. In that case, you cannot use the VCI VisualSpeller feature. You can use it if you have a Borland product on your computer."
ElseIf List1.Text = "Initiate A New Email" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Tools menu and click New Email. A new message window shall then be displayed."
ElseIf List1.Text = "View In Microsoft Word" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Tools menu, highlight view and wait a second for a new menu to appear. Click View In Microsoft Word."
ElseIf List1.Text = "Changing The View" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Window menu and click Cascade to cascade windows, Horizontal to horizontally align the windows. Vertical to align the windows vertically and Arrange Icons to arrange the icons."
ElseIf List1.Text = "Working With Help" Then
    frmHelp.Show
    frmHelp.RichTextBox1.Text = ""
    frmHelp.RichTextBox1.Text = "Go to the Help menu and it will list all different types of help. It will also list program credits as well as help topics."
End If
End Sub

