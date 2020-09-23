VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditor 
   Caption         =   "HTML Editor"
   ClientHeight    =   8070
   ClientLeft      =   3360
   ClientTop       =   2040
   ClientWidth     =   8445
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   8445
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdlDialogBox 
      Left            =   3120
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser wbWYSIWYG 
      Height          =   1935
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   6015
      ExtentX         =   10610
      ExtentY         =   3413
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin RichTextLib.RichTextBox rtbEditor 
      Height          =   1815
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3201
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmEditor.frx":014A
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2400
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":01CC
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":02DE
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":03F0
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":054A
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":06A4
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":07B6
            Key             =   "Picture"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0C08
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0D1A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0E2C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0F3E
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":1050
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":1162
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":1274
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":1386
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":1498
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":15AA
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":16BC
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":17CE
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":18E0
            Key             =   "Justify"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   7800
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5689
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "6/03/02"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "5:01 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbTopBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Index           =   1
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******IDENTIFICATION BLOCK*******
'
'PROGRAMMERS: ZAC ANDREW, KEVIN LAWRENCE, STEVEN PEARCE
'PROJECT NAME: HTMLVILLE
'COURSE CODE: TIK 200-C
'TEACHER: MRS. SINTZEL
'ROOM: 208
'
'*******SPECIFICATION*******
'
'THIS FORM IS AN HTML EDITOR THAT ALLOWS THE USER TO TEST THEIR CODE THAT
'THEY HAVE LEARNED FROM THE LESSONS. THE USER ENTERS HTML CODE INTO THE
'TEXTBOX, AND IT AUTOMATICALLY UPDATES AN HTML FILE WHICH SHOWS WHAT THE
'CODE LOOKS LIKE. THIS EDITOR HAS FEATURES SUCH AS SAVE, OPEN, NEW, PRINT
'ETC.
'
'*******PROGRAM CODE*******

Option Explicit
Dim strDocName As String
Dim DocChanged As String

Private Sub Form_Load()
    wbWYSIWYG.Navigate "about:blank" 'SHOW A BLANK FILE
    
    strDocName = " (Untitled)"
    Me.Caption = App.Title & strDocName 'SHOW FILE NAME IN THE FORM TITLE
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 'SHOW ABOUT FORM
End Sub

Private Sub mnuCopy_Click()
    CopytoClipBoard 'PERFORM FUNCTIONS
    ChangeMenus
End Sub

Public Sub ChangeMenus()

    ' Makes the menus and toolbar context sensative.
mnuSave.Enabled = DocChanged
mnuSaveAs.Enabled = DocChanged
mnuCopy.Enabled = False
mnuCut.Enabled = False
mnuPaste.Enabled = False
tbTopBar.Buttons("Save").Enabled = DocChanged
    
If rtbEditor.SelLength > 0 Then
    mnuCut.Enabled = True
    mnuCopy.Enabled = True
    tbTopBar.Buttons("Cut").Enabled = True
    tbTopBar.Buttons("Copy").Enabled = True

Else
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    tbTopBar.Buttons("Cut").Enabled = False
    tbTopBar.Buttons("Copy").Enabled = False
    
End If

If Clipboard.GetFormat(vbCFText) Then
    mnuPaste.Enabled = True
    tbTopBar.Buttons("Paste").Enabled = True
Else
    mnuPaste.Enabled = False
    tbTopBar.Buttons("Paste").Enabled = True
End If

End Sub
   
Private Sub mnuCut_Click()
    CopytoClipBoard 'PERFORM FUNCTIONS
    ChangeMenus
End Sub

Private Sub mnuDelete_Click()
ChangeMenus 'UPDATE MENUS
End Sub

Private Sub mnuExit_Click()
    Unload Me 'EXITS PROGRAM
End Sub

Private Sub CopytoClipBoard()
    Clipboard.SetText rtbEditor.SelText 'COPY TEXT TO CLIPBOARD
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If DocChanged Then 'IF THE DOCUMENT HAS BEEN CHANGED SINCE ITS LAST SAVE
    
    Select Case MsgBox("The file has changed." & vbCr & vbCr & _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, frmEditor.Caption) 'IF NOT SAVED, GIVES USER CHANGE TO SAVE
    
    Case vbYes
        mnuSave_Click 'IF YES, SAVE DOCUMENT
    Case vbNo
        Unload frmEditor 'IF NO, CLOSE APPLICATION
    Case vbCancel
        Cancel = True 'IF CANCEL, KEEP APPLICATION OPEN
    
    End Select

End If

End Sub

Private Sub mnuNew_Click(Index As Integer)
Dim intCancel As Integer

On Error Resume Next
If DocChanged = False Then
    rtbEditor.Text = "" 'CLEAR TEXT FOR A NEW FILE
Else
    Select Case MsgBox("The file has changed." & vbCr & vbCr & _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, frmEditor.Caption) 'IF NOT SAVED, GIVES USER CHANGE TO SAVE
    
    Case vbYes
        mnuSave_Click 'IF YES, SAVE DOCUMENT
    Case vbNo
        rtbEditor.Text = "" 'IF NO, CLOSE APPLICATION
    Case vbCancel
        intCancel = True 'IF CANCEL, KEEP APPLICATION OPEN
    
    End Select
End If
End Sub

Private Sub mnuOpen_Click()
Dim Cancel As Boolean
On Error GoTo errorhandler
Cancel = False

cdlDialogBox.Filter = "HTML Files (*.html)|*.html|Text Files (*.txt)|*.txt|All Files|*.*" 'TYPES OF FILES TO BE OPENED
cdlDialogBox.CancelError = True
cdlDialogBox.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
cdlDialogBox.ShowOpen

If Not Cancel Then
    If UCase(Right(cdlDialogBox.FileName, 3)) = "RTF" Then
        rtbEditor.LoadFile cdlDialogBox.FileName, rtfRTF
    Else
        rtbEditor.LoadFile cdlDialogBox.FileName, rtfText
    End If
        rtbEditor.FileName = cdlDialogBox.FileName
        strDocName = rtbEditor.FileName
        Me.Caption = App.Title & " " & strDocName 'PUT DOCUMENT NAME IN APPLICATION TITLE
        DocChanged = False 'DOCUMENT HAS NOT BEEN CHANGED
End If
Exit Sub

errorhandler: 'INCASE OF AN ERROR
If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
End If
End
End Sub

Private Sub mnuPaste_Click()
Dim Text As String
Dim ClipboardText As String
Dim SelStart As Long
    
If Clipboard.GetFormat(vbCFText) Then 'CHECK IF THERE IS TEXT ON CLIPBOARD
    
    If rtbEditor.SelLength > 0 Then 'REPLACE SELECTED TEXT
        Exit Sub
    End If
    
    Text = rtbEditor.Text 'MOVE TEXT TO VARIABLE
    SelStart = rtbEditor.SelStart
    ClipboardText = Clipboard.GetText
    
    rtbEditor.Text = Left(Text, SelStart) & _
            ClipboardText & Right(Text, Len(Text) - SelStart) 'REPLACE SELECTED TEXT WITH STRING
    
    rtbEditor.SelStart = SelStart 'PUT CURSOR BACK

Else
    ChangeMenus 'PERFORM FUNCTION
End If
End Sub

Private Sub mnuPrint_Click()
Dim bcancel As Boolean
Dim ncopy As Integer
On Error GoTo errorhandler

bcancel = False

cdlDialogBox.Flags = cdlPDHidePrintToFile Or _
        cdlPDNoSelection Or cdlPDNoPageNums _
        Or cdlPDCollate

cdlDialogBox.CancelError = True
cdlDialogBox.PrinterDefault = True
cdlDialogBox.Copies = 1
cdlDialogBox.ShowPrinter

If bcancel = False Then
    PrintRTF rtbEditor, 1440, 1440, 1440, 1440
    For ncopy = 1 To cdlDialogBox.Copies
    Next ncopy
End If

Exit Sub

errorhandler:
If Err.Number = cdlCancel Then
bcancel = True
Resume Next
End If
End Sub

Private Sub mnuSave_Click()
    If strDocName = " (Untitled)" Then
        mnuSaveAs_Click
    Else
        If UCase(Right(rtbEditor.FileName, 3)) = "RTF" Then
            rtbEditor.SaveFile rtbEditor.FileName, rtfRTF
        Else
            rtbEditor.SaveFile rtbEditor.FileName, rtfText
        End If
        
        DocChanged = False
    End If
End Sub

Private Sub mnuSaveAs_Click()
    Dim Cancel As Boolean
    On Error GoTo errorhandler
    Cancel = False

    cdlDialogBox.DefaultExt = ".html"
    cdlDialogBox.Filter = "HTML Files (*.html)|*.html|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    cdlDialogBox.CancelError = True
    cdlDialogBox.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt

    cdlDialogBox.ShowSave

    If Not Cancel Then
        If UCase(Right(cdlDialogBox.FileName, 3)) = "RTF" Then
            rtbEditor.SaveFile cdlDialogBox.FileName, rtfRTF
        Else
            rtbEditor.SaveFile cdlDialogBox.FileName, rtfText
        End If
    rtbEditor.FileName = cdlDialogBox.FileName
    strDocName = cdlDialogBox.FileName
    Me.Caption = App.Title & " " & strDocName
    DocChanged = False
End If

Exit Sub

errorhandler:
If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
End If
End Sub


Private Sub mnuStatusBar_Click()
    ' Shows or hides the status bar as needed.
    ' And makes the menu context sensative.
mnuStatusBar.Checked = Not mnuStatusBar.Checked
sbStatusBar.Visible = mnuStatusBar.Checked
End Sub

Private Sub mnuToolbar_Click()
    ' Shows or hides the toolbar as needed.
    ' And makes the menu context sensative.
mnuToolbar.Checked = Not mnuToolbar.Checked
tbTopBar.Visible = mnuToolbar.Checked

    ' This resizes the richtextbox depending on the state of the toolbar.

End Sub

Private Sub rtbEditor_Change()
    DoEvents
        Open "C:\temporary.html" For Output As #1: Print #1, rtbEditor.Text: Close #1 'CREATE A TEMPORARY FILE
    DoEvents
        wbWYSIWYG.Navigate "C:\temporary.html" 'SHOW THE TEMPORARY FILE
        
    DocChanged = True
End Sub

Private Sub Form_Resize()
    wbWYSIWYG.Top = 500 'CHANGE THE SIZES WHEN THE APPLICATION IS RESIZED
    wbWYSIWYG.Left = 40
    wbWYSIWYG.Width = Me.Width - 300
    wbWYSIWYG.Height = Me.Height / 2 - 300

    rtbEditor.Top = Me.Height / 2 + 200
    rtbEditor.Left = 40
    rtbEditor.Width = Me.Width - 300
    rtbEditor.Height = Me.Height / 2 - 1200
End Sub


Private Sub tbTopBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key 'GIVE FUNCTIONS FOR TOOLBAR
            
        Case "Open"
            mnuOpen_Click
        
        Case "Save"
            mnuSave_Click
        
        Case "Print"
            mnuPrint_Click
            
        Case "Cut"
            mnuCut_Click
        
        Case "Copy"
            mnuCopy_Click
            
        Case "Paste"
            mnuPaste_Click
    End Select
End Sub
