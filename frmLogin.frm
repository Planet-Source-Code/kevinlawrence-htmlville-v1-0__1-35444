VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Welcome to HTML Ville please Login"
   ClientHeight    =   2925
   ClientLeft      =   4140
   ClientTop       =   4230
   ClientWidth     =   7095
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7095
   Begin VB.CommandButton cmdDone 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.Data dbDatabar 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\HTMLville\db.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   1  'EOF
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Users"
      Top             =   2760
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      DataSource      =   "databar"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox txtUsername 
      DataField       =   "User"
      DataSource      =   "databar"
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdNewUser 
      Caption         =   "New User"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image imgHTMLville 
      Height          =   2970
      Left            =   1440
      Picture         =   "frmLogin.frx":014A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Username: (case sensative)"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H80000005&
      Caption         =   "Enter Password:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
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
'THIS FORM ALLOWS THE USER TO LOGIN TO HTMLVILLE UNDER HIS/HER ACCOUNT.
'THE FORM ACCESS' AN ACCESS DATABASE, AND IF THE LOGIN/USERNAME SUBMITTED
'MATCH A RECORD IN THE DATABASE, THE USER WILL BE LOGGED IN.
'
'*******PROGRAM CODE*******

Option Explicit
Dim strUsername As String, strPassword As String
Private dbWorkspace As Workspace
Private dbDatabase As Database
Private dbTable As Recordset
Dim dbTableDef As TableDef
Dim dbUserName As Field
Dim dbPassword As Field
Dim dbLevel As Field

Private Sub cmdDone_Click()
    Unload frmLogin
End Sub

Private Sub cmdEnter_Click()
    Dim frmLoginDB As Database 'DECLARE VARIABLES
    Dim frmLoginRecordSet As Recordset
    
    Set frmLoginDB = OpenDatabase(App.Path & "\db.mdb", True, False, ";pwd=" & "HTMLville") 'CONNECT TO DATABASE
    Set frmLoginRecordSet = frmLoginDB.OpenRecordset("Users") 'OPEN THE USERS TABLE
    
    Do While Not frmLoginRecordSet.EOF 'DO WHILE SERACHING THE DATABASE
    If frmLoginRecordSet.Fields("User") = (txtUsername.Text) And _
    frmLoginRecordSet.Fields("Password") = (txtPassword.Text) Then

    frmMain.Show 'SHOW THE MAIN MENU
    
    Unload Me 'CLOSE THE LOGIN
    
    Exit Sub
    
    Else
    frmLoginRecordSet.MoveNext
    End If
    Loop
    txtPassword.Text = "" 'CLEAR THE TEXTBOX
    MsgBox "Invalid Username or Password. Please Try Again.", vbOKOnly + vbCritical + vbSystemModal, "HTMLville: Error" 'SHOW ERROR MESSAGE
End Sub

Private Sub cmdNewUser_Click()
    frmNewUser.Show 'OPEN NEW USER FORM
End Sub

Private Sub Form_Load()
    strUsername = txtUsername.Text
    strPassword = txtPassword.Text
    txtUsername.Text = ""
    txtPassword.Text = ""
End Sub
