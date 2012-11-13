VERSION 5.00
Begin VB.Form Frm0 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      Picture         =   "Frm0.frx":0000
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   16
      Top             =   5760
      Width           =   735
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      Picture         =   "Frm0.frx":182A
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   15
      Top             =   6480
      Width           =   735
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      Picture         =   "Frm0.frx":3054
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   14
      Top             =   5040
      Width           =   735
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      Picture         =   "Frm0.frx":487E
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   13
      Top             =   4320
      Width           =   735
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      Picture         =   "Frm0.frx":60A8
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   12
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      Picture         =   "Frm0.frx":78D2
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   11
      Top             =   3600
      Width           =   735
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      Picture         =   "Frm0.frx":90FC
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   8
      Top             =   2160
      Width           =   735
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         Picture         =   "Frm0.frx":A926
         ScaleHeight     =   705
         ScaleWidth      =   705
         TabIndex        =   9
         Top             =   0
         Width           =   735
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            FontTransparent =   0   'False
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   0
            Picture         =   "Frm0.frx":C150
            ScaleHeight     =   705
            ScaleWidth      =   705
            TabIndex        =   10
            Top             =   0
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      Picture         =   "Frm0.frx":D97A
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
      Left            =   240
      Picture         =   "Frm0.frx":F1A4
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
      Left            =   240
      Picture         =   "Frm0.frx":109CE
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
      Picture         =   "Frm0.frx":121F8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
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
      Left            =   2640
      Picture         =   "Frm0.frx":138C6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I'm sorry!"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Clearly you're a white sox fan... There's no hope! you got 0 right!"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2880
      Width           =   6255
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   2895
      Left            =   1440
      TabIndex        =   4
      Top             =   2640
      Width           =   6735
   End
End
Attribute VB_Name = "Frm0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: Frm0
'Project By: Stephanie Arel
'Date Written: 3/16/2009
'The purpose of this form is to tell the user that they scored 0/5 on the trivia.
Option Explicit


Private Sub Command1_Click()
'Takes the user back to the main form
Frm0.Hide
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
