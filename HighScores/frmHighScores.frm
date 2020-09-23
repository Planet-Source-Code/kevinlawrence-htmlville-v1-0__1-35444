VERSION 5.00
Begin VB.Form frmHighScores 
   BackColor       =   &H80000005&
   Caption         =   "High Scores"
   ClientHeight    =   3750
   ClientLeft      =   5115
   ClientTop       =   3660
   ClientWidth     =   4680
   Icon            =   "frmHighScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4680
   Begin VB.Label lblScoreTitle 
      BackColor       =   &H80000005&
      Caption         =   "Score"
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
      Left            =   3360
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblUsernameTitle 
      BackColor       =   &H80000005&
      Caption         =   "Name"
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
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblScore 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblScore 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblScore 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblScore 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblScore 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmHighScores"
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
'THIS FORM DISPLAYS THE HIGH SCORES FROM THE QUIZZES. IT GETS THE
'INFORMATION FROM A TEXT FILE AND DISPLAYS IT IN ARRAYS OF LABELS.
'
'*******PROGRAM CODE*******

Option Explicit
Dim db As Database
Dim rec As Recordset
Dim i As Integer

Private Sub Form_Activate()
Set db = OpenDatabase("db.mdb", True, False, ";pwd=HTMLville")
Set rec = db.OpenRecordset("select * from Quiz order by Score")
rec.MoveLast
    For i = 0 To 4
        lblUsername(i) = rec!UserName
        lblScore(i) = rec!score
        rec.MovePrevious
    Next
End Sub
