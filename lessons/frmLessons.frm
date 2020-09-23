VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmLessons 
   Caption         =   "Lessons"
   ClientHeight    =   3345
   ClientLeft      =   5115
   ClientTop       =   3840
   ClientWidth     =   5100
   Icon            =   "frmLessons.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5100
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdLesson3 
      Caption         =   "Lesson 3"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdLesson2 
      Caption         =   "Lesson 2"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdLesson1 
      Caption         =   "Lesson 1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser wbBrowser 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      ExtentX         =   8493
      ExtentY         =   4683
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
End
Attribute VB_Name = "frmLessons"
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
'THIS FORM ALLOWS THE USER TO VIEW DIFFERENT HTML LESSONS. THE BUTTONS
'DIRECT THE BROWSER TO DIFFERENT HTML FILES WHICH ARE IN THE CODELIBRARY
'FOLDER OF THE PROGRAM FILES.
'
'*******PROGRAM CODE*******

Option Explicit

Private Sub cmdBack_Click()
On Error Resume Next
    wbBrowser.GoBack
End Sub

Private Sub cmdLesson1_Click()
    wbBrowser.Navigate ("C:\Program Files\HTMLville\Lessons\Lesson1.html") 'NAVIGATE TO LESSON 1
End Sub

Private Sub cmdLesson2_Click()
    wbBrowser.Navigate ("C:\Program Files\HTMLville\Lessons\Lesson2.html") 'NAVIGATE TO LESSON 2
    cmdBack.Enabled = True 'SHOW BACK BUTTON
End Sub

Private Sub cmdLesson3_Click()
    wbBrowser.Navigate ("C:\Program Files\HTMLville\Lessons\Lesson3.html") 'NAVIGATE TO LESSON 3
    cmdBack.Enabled = True 'SHOW BACK BUTTON
End Sub

Private Sub Form_Resize()
    wbBrowser.Top = 600 'CHANGE THE SIZES WHEN THE APPLICATION IS RESIZED
    wbBrowser.Left = 40
    wbBrowser.Width = Me.Width - 300
    wbBrowser.Height = Me.Height - 1100
End Sub

Private Sub Form_Load()
    wbBrowser.Navigate ("C:\Program Files\HTMLville\Lessons\Lesson1.html") 'NAVIGATE TO LESSON 1
End Sub
