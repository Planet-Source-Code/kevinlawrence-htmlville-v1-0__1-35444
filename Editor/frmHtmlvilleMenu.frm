VERSION 5.00
Begin VB.Form frmHtmlvilleMenu 
   BackColor       =   &H8000000E&
   Caption         =   "Htmlville Menu"
   ClientHeight    =   7020
   ClientLeft      =   840
   ClientTop       =   1020
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10170
   Begin VB.Label lblMessage 
      BackColor       =   &H8000000E&
      Caption         =   "*Click where you want to go*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblSteven 
      BackColor       =   &H8000000E&
      Caption         =   "Steven Pearce"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblKevin 
      BackColor       =   &H8000000E&
      Caption         =   "Kevin Lawrence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblZac 
      BackColor       =   &H8000000E&
      Caption         =   "Zac Andrew"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblDevelopers 
      BackColor       =   &H8000000E&
      Caption         =   "Developers:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image imgClasses 
      Height          =   1980
      Left            =   0
      Picture         =   "frmHtmlvilleMenu.frx":0000
      Top             =   4800
      Width           =   2475
   End
   Begin VB.Image imgDataBank 
      Height          =   2145
      Left            =   6720
      Picture         =   "frmHtmlvilleMenu.frx":10002
      Top             =   4440
      Width           =   3360
   End
   Begin VB.Image imgHtmlvilleLogo 
      Height          =   2805
      Left            =   2880
      Picture         =   "frmHtmlvilleMenu.frx":277A4
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Image imgTestingOffice 
      Height          =   2325
      Left            =   6480
      Picture         =   "frmHtmlvilleMenu.frx":4ABD2
      Top             =   1320
      Width           =   2940
   End
   Begin VB.Image imgHtmlEditor 
      Height          =   1785
      Left            =   240
      Picture         =   "frmHtmlvilleMenu.frx":61018
      Top             =   2040
      Width           =   2805
   End
   Begin VB.Image imgHighScore 
      Height          =   2340
      Left            =   3120
      Picture         =   "frmHtmlvilleMenu.frx":71686
      Top             =   0
      Width           =   3570
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmHtmlvilleMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

