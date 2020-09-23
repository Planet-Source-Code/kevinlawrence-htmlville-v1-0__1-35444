VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "HTMLville"
   ClientHeight    =   5880
   ClientLeft      =   3555
   ClientTop       =   3000
   ClientWidth     =   8130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8130
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash sfMusicPlayer 
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   5160
      Width           =   1575
      _cx             =   2778
      _cy             =   1085
      FlashVars       =   ""
      Movie           =   "C:\Program Files\HTMLville\music.swf"
      Src             =   "C:\Program Files\HTMLville\music.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Image imgBank 
      Height          =   2025
      Left            =   0
      Picture         =   "frmMain.frx":014A
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2760
   End
   Begin VB.Image imgEditor 
      Height          =   1665
      Left            =   0
      Picture         =   "frmMain.frx":178EC
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2325
   End
   Begin VB.Image imgTest 
      Height          =   2085
      Left            =   5400
      Picture         =   "frmMain.frx":27F5A
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2580
   End
   Begin VB.Image imgClasses 
      Height          =   1980
      Left            =   5640
      Picture         =   "frmMain.frx":3E3A0
      Top             =   3840
      Width           =   2475
   End
   Begin VB.Image imgScores 
      Height          =   1740
      Left            =   2640
      Picture         =   "frmMain.frx":4E3A2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2850
   End
   Begin VB.Image imgVille 
      Height          =   2205
      Left            =   2760
      Picture         =   "frmMain.frx":69834
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuChangeUser 
         Caption         =   "&ChangeUser"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
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
'THIS FORM IS THE MAIN FORM OF THE APPLICATION. IT IS THE MENU THAT
'ALLOWS THE USER TO ACCESS THE DIFFERENT FEATURES OF HTMLVILLE. IT ALSO
'HAS A MUSIC BOX THAT CAN BE TURNED ON OR OFF, WHICH IS MADE ON
'MACROMEDIA FLASH.
'
'*******PROGRAM CODE*******

Option Explicit


Private Sub imgBank_Click()
    frmCodeLibrary.Show 'SHOW THE CODE LIBRARY FORM
End Sub

Private Sub imgClasses_Click()
    frmLessons.Show 'SHOW THE LESSONS FORM
End Sub

Private Sub imgEditor_Click()
    frmEditor.Show 'SHOW THE EDITOR FORM
End Sub

Private Sub imgScores_Click()
    frmHighScores.Show 'SHOW THE QUIZ FORM WITH HIGH SCORES
End Sub

Private Sub imgTest_Click()
    frmQuizMenu.Show 'SHOW THE QUIZ MENU FORM
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 'SHOW THE ABOUT FORM
End Sub

Private Sub mnuChangeUser_Click()
    frmLogin.Show 'SHOW THE LOGIN FORM
    Unload frmMain 'EXIT FORM
End Sub

Private Sub mnuExit_Click()
    Unload frmMain 'EXIT APPLICATION
End Sub
