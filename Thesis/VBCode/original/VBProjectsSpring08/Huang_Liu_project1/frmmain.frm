VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Main Page"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   LinkTopic       =   "Form2"
   Picture         =   "frmmain.frx":0000
   ScaleHeight     =   7425
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log in"
      Height          =   975
      Left            =   6120
      Picture         =   "frmmain.frx":C774
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdmsg 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Message Board"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      Picture         =   "frmmain.frx":CB9A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdshop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9000
      Picture         =   "frmmain.frx":141FC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdreadme 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Read Me"
      Height          =   975
      Left            =   7200
      Picture         =   "frmmain.frx":19E52
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdfriends 
      BackColor       =   &H00FFFFFF&
      Caption         =   "See Kitty's Friends"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   480
      Picture         =   "frmmain.frx":1A252
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdfile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "See the shopping catalog"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6960
      Picture         =   "frmmain.frx":251BC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdquestion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Characters' Challenge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2640
      Picture         =   "frmmain.frx":2876E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      Picture         =   "frmmain.frx":31628
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdgame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kitty's guess number game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4800
      Picture         =   "frmmain.frx":36BBA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblname 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please log in first."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdfile_Click()
frmfile.Visible = True
frmmain.Visible = False
End Sub

Private Sub cmdfriends_Click()
frmfriend.Visible = True
frmmain.Visible = False
End Sub

Private Sub cmdgame1_Click()
frmmain.Visible = False
frmgame.Visible = True
End Sub

Private Sub cmdlogin_Click()
Customer = InputBox("Please Enter Your User Name", "Name")
Dates = InputBox("Please Enter The Date", "Date")
lblname.Caption = Customer & "     " & Dates
End Sub

Private Sub cmdmsg_Click()
frmdisplay.Show
frmmain.Hide
End Sub

Private Sub cmdquestion_Click()
frmmain.Visible = False
frmquestion.Visible = True
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreadme_Click()
frmmain.Visible = False
frmreadme.Visible = True
End Sub

Private Sub cmdshop_Click()
frmmain.Visible = False
frmshop.Visible = True
End Sub
