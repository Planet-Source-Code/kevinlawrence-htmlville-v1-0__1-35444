VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmCodeLibrary 
   Caption         =   "Code Library"
   ClientHeight    =   4920
   ClientLeft      =   2970
   ClientTop       =   3270
   ClientWidth     =   9270
   Icon            =   "frmCodeLibrary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9270
   Begin SHDocVwCtl.WebBrowser wbBrowser 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   9015
      ExtentX         =   15901
      ExtentY         =   7435
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
      Location        =   ""
   End
   Begin VB.CommandButton cmdHTML 
      Caption         =   "HTML"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdJavascript 
      Caption         =   "Javascript"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmCodeLibrary"
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
'THIS FORM ALLOWS THE USER TO VIEW DIFFERENT CODE SNIPPETS IN A BROWSER
'THAT GOES TO HTML FILES WHICH ARE IN THE CODELIBRARY FOLDER OF THE
'PROGRAM FILES.
'
'*******PROGRAM CODE*******

Option Explicit

Private Sub cmdHTML_Click()
    wbBrowser.Navigate ("C:\Program Files\HTMLville\CodeLibrary\html\default.html") 'NAVIGATE TO HTML PAGE
End Sub

Private Sub cmdJavascript_Click()
    wbBrowser.Navigate ("C:\Program Files\HTMLville\CodeLibrary\javascript\default.html") 'NAVIGATE TO JAVASCRIPT PAGE
End Sub

Private Sub Form_Load()
    wbBrowser.Navigate ("C:\Program Files\HTMLville\CodeLibrary\javascript\default.html") 'NAVIGATE TO JAVASCRIPT PAGE
End Sub
