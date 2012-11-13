VERSION 5.00
Begin VB.Form Frm5 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   DrawMode        =   16  'Merge Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":0000
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   14
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":182A
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   13
      Top             =   5760
      Width           =   735
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":3054
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   12
      Top             =   6480
      Width           =   735
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":487E
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   11
      Top             =   5040
      Width           =   735
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":60A8
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   10
      Top             =   4320
      Width           =   735
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":78D2
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":90FC
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":A926
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":C150
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "Frm5.frx":D97A
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Picture         =   "Frm5.frx":F1A4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Picture         =   "Frm5.frx":10872
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " CONGRATS!"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   48
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   " You are a trivia WIZARD! You got all 5 questions correct!! :)"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   26.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1560
      TabIndex        =   0
      Top             =   2640
      Width           =   6255
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   2655
      Left            =   1320
      TabIndex        =   2
      Top             =   2280
      Width           =   6735
   End
End
Attribute VB_Name = "Frm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: Frm5
'Project By: Stephanie Arel
'Date Written: 3/16/2009
'The purpose of this form is to tell the user that they scored 5/5 on the trivia!
Option Explicit

Private Sub Command1_Click()
'Takes the user back to the main menu.
Frm5.Hide
FrmMain.Show
End Sub

Private Sub Command2_Click()
'Ends the program completely.
End
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
