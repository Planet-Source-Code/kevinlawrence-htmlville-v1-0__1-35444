VERSION 5.00
Begin VB.Form frmQuizMenu 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quizzes"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6990
   Icon            =   "frmQuizMenu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuiz3 
      Caption         =   "Quiz 3"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuiz2 
      Caption         =   "Quiz 2"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuiz1 
      Caption         =   "Quiz 1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblQuiz3 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Based on skills learned in lesson 3"
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
      Height          =   855
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblQuiz2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Based on skills learned in lesson 2"
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
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblQuiz1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Based on skills learned in lesson 1"
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
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmQuizMenu"
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
'THIS FORM IS A MENU FOR THE QUIZ AND ALLOWS THE USER TO SELECT WHICH
'QUIZ HE/SHE WANTS TO TAKE. THE FORM LOADS THE QUIZ FILE FOR THE QUIZ
'FORM, WHICH HOLDS THE QUESTIONS AND ANSWERS.
'
'*******PROGRAM CODE*******

Option Explicit

Private Sub cmdQuiz1_Click()
    file = "Quiz1.txt"
    frmQuiz.Show
    Unload frmQuizMenu
End Sub

Private Sub cmdQuiz2_Click()
    file = "Quiz2.txt"
    frmQuiz.Show
    Unload frmQuizMenu
End Sub

Private Sub cmdQuiz3_Click()
    file = "Quiz3.txt"
    frmQuiz.Show
    Unload frmQuizMenu
End Sub
