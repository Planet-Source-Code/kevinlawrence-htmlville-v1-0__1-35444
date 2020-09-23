VERSION 5.00
Begin VB.Form frmNewUser 
   BackColor       =   &H80000005&
   Caption         =   "Creating new account"
   ClientHeight    =   1335
   ClientLeft      =   6285
   ClientTop       =   4620
   ClientWidth     =   3375
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3375
   Begin VB.Data dbDatabar 
      Caption         =   "New User"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\HTMLville\db.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Users"
      Top             =   1200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtNewUser 
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtNewPass 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Account"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblNewPassword 
      BackColor       =   &H80000005&
      Caption         =   "Enter Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblNewUser 
      BackColor       =   &H80000005&
      Caption         =   "User Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmNewUser"
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
'THIS FORM ALLOWS THE USER TO CREATE A NEW ACCOUNT TO BE USED. IT WILL
'ASK FOR THE DESIRED USERNAME/PASSWORD AND WILL SUBMIT IT TO THE DATABASE.
'ONCE IT IS SUBMITTED, IT SHOWS A SUCCESS MESSAGE.
'
'*******PROGRAM CODE*******

Option Explicit
Private dbWorkspace As Workspace
Private dbDatabase As Database
Private dbTable As Recordset

Private Sub cmdCancel_Click()
    Unload Me 'EXIT PROGRAM
End Sub

Private Sub cmdCreate_Click()
    If txtNewUser = "" Or txtNewPass = "" Then 'IF FIELDS ARE EMPTY
    Exit Sub
    End If
    
    dbTable.MoveLast
    dbTable.AddNew
    dbTable!User = txtNewUser.Text 'ADD TEXT FROM USER TEXTBOX TO USER COLUMN
    dbTable!Password = txtNewPass.Text 'ADD TEXT FROM PASSWORD TEXTBOX TO PASSWORD COLUMN
    dbTable.Update
    MsgBox "An account for " + txtNewUser.Text + " has been created successfully!", vbOKOnly, "HTMLville: Add User" 'TELL USER SUCCESS
    txtNewUser.Text = "" 'CLEAR TEXTBOXES
    txtNewPass.Text = ""
    Unload frmNewUser
End Sub

Private Sub Form_Load()
    Set dbWorkspace = DBEngine.Workspaces(0)
    Set dbDatabase = dbWorkspace.OpenDatabase(App.Path & "\db.mdb", True, False, ";pwd=" & "HTMLville") 'CONNECT TO DATABASE
    Set dbTable = dbDatabase.OpenRecordset("Users", dbOpenTable) 'OPEN USERS TABLE
End Sub
