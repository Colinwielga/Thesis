VERSION 5.00
Begin VB.Form Frm3 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":0000
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":182A
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   13
      Top             =   5880
      Width           =   735
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":3054
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   12
      Top             =   6600
      Width           =   735
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":487E
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   11
      Top             =   5160
      Width           =   735
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":60A8
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   10
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":78D2
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   9
      Top             =   3000
      Width           =   735
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":90FC
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   8
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":A926
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":C150
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      Picture         =   "Frm3.frx":D97A
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   120
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
      Left            =   5160
      Picture         =   "Frm3.frx":F1A4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
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
      Left            =   2760
      Picture         =   "Frm3.frx":10872
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Not too Bad!"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   " You are pretty Good! You could use a little bit of work though! (3 right)"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   26.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1800
      TabIndex        =   2
      Top             =   2760
      Width           =   6255
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   2895
      Left            =   1560
      TabIndex        =   4
      Top             =   2520
      Width           =   6735
   End
End
Attribute VB_Name = "Frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: Frm3
'Project By: Stephanie Arel
'Date Written: 3/16/2009
'The purpose of this form is to tell the user that they scored 3/5 on the trivia.
Option Explicit

Private Sub Command1_Click()
'Takes user back to the main menu.
Frm3.Hide
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
