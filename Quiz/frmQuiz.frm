VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTMLville Quiz"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmQuiz.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraHighScores 
      BackColor       =   &H80000005&
      Height          =   2775
      Left            =   7080
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Label lblUsername 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   26
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   25
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   24
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   23
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   22
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label lblScore 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblScore 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblScore 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblScore 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblScore 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   17
         Top             =   2160
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "&Menu"
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.Frame fraQuiz 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.Timer tmrInterval 
         Interval        =   1
         Left            =   0
         Top             =   120
      End
      Begin VB.OptionButton optAnswer 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   2895
      End
      Begin VB.OptionButton optAnswer 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   1
         Left            =   3840
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton optAnswer 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton optAnswer 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   3
         Left            =   3840
         TabIndex        =   5
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label lblD 
         BackColor       =   &H80000005&
         Caption         =   "d.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lblB 
         BackColor       =   &H80000005&
         Caption         =   "b.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblC 
         BackColor       =   &H80000005&
         Caption         =   "c.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lblA 
         BackColor       =   &H80000005&
         Caption         =   "a.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblQuestion 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame fraStats 
      BackColor       =   &H80000005&
      Caption         =   "Your Results"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   7080
      TabIndex        =   10
      Top             =   0
      Width           =   6855
      Begin VB.Label lblAnswered 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   5895
      End
      Begin VB.Label lblAnswersWrong 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   5895
      End
      Begin VB.Label lblAnswersRight 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmQuiz"
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
'THIS FORM QUIZZES THE USER ON THE 3 LESSONS FROM THE LESSONS FORM. ONCE
'THE USER COMPLETES THE QUIZ, HIS/HER SCORE IS SUBMITTED TO THE HIGH
'SCORES TEXT FILE.
'*******PROGRAM CODE*******

Dim hin(0 To 4) As String
Dim hisc(0 To 4) As String
Dim score As Integer
Dim u As Integer
Dim d As Integer
Dim t As Integer
Dim c As Integer
Dim r As Integer
Dim intFor As Integer

Dim record(4) As question, pos As Integer
Dim intX As Integer
Dim ends As Boolean
Dim intQno As Integer
Dim intCorrect As Integer
Dim intIncorrect As Integer
Dim yesno(4) As String
Dim qano(4) As Boolean
Private dbWorkspace As Workspace
Private dbDatabase As Database
Private dbTable As Recordset

Private Function openfile()
    Open App.Path & "\" & file For Random As 1 Len = Len(record(4))
        pos = pos + 1
        Get #1, pos, record(pos - 1)
        fraQuiz.Caption = "Question " & record(pos - 1).intQno
        lblQuestion.Caption = record(pos - 1).ques
        optAnswer(0).Caption = record(pos - 1).ans1
        optAnswer(1).Caption = record(pos - 1).ans2
        optAnswer(2).Caption = record(pos - 1).ans3
        optAnswer(3).Caption = record(pos - 1).ans4
    Close 1
End Function
Private Function clear()
    optAnswer(0).Value = False
    optAnswer(1).Value = False
    optAnswer(2).Value = False
    optAnswer(3).Value = False
End Function

Private Sub cmdNext_Click()
    
    If (optAnswer(0).Value = False And optAnswer(1).Value = False And optAnswer(2).Value = False And optAnswer(3).Value = False) Then
        MsgBox "Please Select A Answer", vbCritical, "PLS"
        Exit Sub
    End If
    If record(pos - 1).intQno = 5 Then
        cmdNext.Enabled = False
        ends = True
    End If
    
    If optAnswer(record(pos - 1).ans).Value = True Then
        MsgBox "You Are right!", vbInformation, "HTML Quiz"
        yesno(pos - 1) = "Right"
        qano(pos - 1) = True
    Else
        MsgBox "You Are wrong.", vbCritical, "HTML Quiz"
        qano(pos - 1) = False
        yesno(pos - 1) = "Wrong"
    End If
    
    If record(pos - 1).intQno = 5 Then
        For i = 0 To 4
            If yesno(i) = "Right" Then
                intCorrect = intCorrect + 1
            Else
                intIncorrect = intIncorrect + 1
            End If
        Next
        For intX = 0 To 4
            If qano(intX) = False Then
                lblAnswered.Caption = lblAnswered.Caption & "Q" & (intX + 1) & ", "
            End If
        Next
        lblAnswersRight.Caption = "You answered " & intCorrect & " questions correctly!"
        lblAnswersWrong.Caption = "You answered " & intIncorrect & " questions incorrectly."
            

            If intCorrect >= Val(lblScore(intFor).Caption) Then
               Do
                 r = r - 1
                 If r <= intFor Then Exit Do
                 lblUsername(r).Caption = lblUsername(r - 1).Caption
                 lblUsername(r - 1).Caption = ""
                 lblScore(r).Caption = lblScore(r - 1).Caption
                 lblScore(r - 1).Caption = ""
               Loop
               Do
                 Dim Name As String
                 Name = InputBox("Enter Your Name(Max. 20 Characters):", "High Score Table Entry", "Unknown")
                 If Len(Name) <= 20 Then
                    Exit Do
                 End If
               Loop
               lblUsername(intFor).Caption = Name
               lblScore(intFor).Caption = intCorrect
               Close #1

    dbTable.MoveLast
    dbTable.AddNew
    dbTable!UserName = lblUsername(c) 'ADD TEXT FROM VARIABLE TO USERNAME COLUMN
    dbTable!score = lblScore(c) 'ADD TEXT FROM VARIABLE TO SCORE COLUMN
    dbTable.Update
               
            End If
        
        If intCorrect = 5 Then
            lblAnswered.Caption = ""
        Else
            lblAnswered.Caption = lblAnswered.Caption & " Wrongly"
        End If
    End If
    
    If record(pos - 1).intQno < 5 Then
        openfile
    End If
    clear
End Sub


Private Sub cmdExit_Click()
    frmQuizMenu.Show 'SHOW QUIZ MENU
    Unload frmQuiz 'UNLOAD FORM
End Sub

Private Sub cmdMenu_Click()
    Unload frmQuiz 'UNLOAD FORM
    frmQuizMenu.Show 'SHOW QUIZ MENU
End Sub
Private Sub Form_Load()
Randomize
   Set dbWorkspace = DBEngine.Workspaces(0)
    Set dbDatabase = dbWorkspace.OpenDatabase(App.Path & "\db.mdb", True, False, ";pwd=" & "HTMLville") 'CONNECT TO DATABASE
    Set dbTable = dbDatabase.OpenRecordset("Quiz", dbOpenTable) 'OPEN QUIZ TABLE
    
    pos = 0
    openfile
    clear
    ends = False
    lblAnswered.Caption = "You Have Answered "
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intQno = 0
    intCorrect = 0
    intIncorrect = 0
End Sub

Private Sub tmrInterval_Timer()
    If ends = True Then
        If fraStats.Left > 120 Then
            fraStats.Left = fraStats.Left - 40
            fraQuiz.Left = fraQuiz.Left - 41
        End If
    End If
End Sub
