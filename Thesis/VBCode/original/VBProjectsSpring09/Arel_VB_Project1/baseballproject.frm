VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6675
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.PictureBox Picture2 
         Height          =   6375
         Left            =   120
         Picture         =   "baseballproject.frx":0000
         ScaleHeight     =   6315
         ScaleWidth      =   8475
         TabIndex        =   1
         Top             =   120
         Width           =   8535
         Begin VB.CommandButton Command5 
            Caption         =   "View Schedule!"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   5880
            Picture         =   "baseballproject.frx":B0686
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   3960
            Width           =   1575
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000009&
            Height          =   2055
            Left            =   240
            ScaleHeight     =   1995
            ScaleWidth      =   7515
            TabIndex        =   3
            Top             =   3840
            Width           =   7575
            Begin VB.CommandButton Command4 
               Caption         =   "Quit"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2760
               Picture         =   "baseballproject.frx":B1D54
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   1440
               Width           =   2175
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Test Your Trivia Knowledge!"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   3840
               Picture         =   "baseballproject.frx":B3422
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Meet the Players!"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1920
               Picture         =   "baseballproject.frx":B4AF0
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   120
               Width           =   1695
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Visit Our Online Hat Store!"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   120
               Picture         =   "baseballproject.frx":B61BE
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   120
               Width           =   1695
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "WELCOME TWIN'S FAN!"
            BeginProperty Font 
               Name            =   "Cooper Black"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   480
            TabIndex        =   2
            Top             =   240
            Width           =   7695
         End
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmMain
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This is the main choosing menu. Here the user can select a variety of options. Each will open a new form.
'The purpose of this project is to allow a Minnesota Twins fan to get to know the team, schedule, test their trivia knowledge and perhaps buy a Twins Hat.

Option Explicit

Private Sub Command1_Click()
'Takes the user to a Hat Store. Here, users can purchase a hat of their choice.
FrmMain.Hide
FrmHats.Show
End Sub

Private Sub Command2_Click()
'Takes the user to a Players menu where they can view their favorite players.
FrmMain.Hide
FrmPlayers.Show
End Sub

Private Sub Command3_Click()
'Takes the User to a Trivia Program.
FrmMain.Hide
FrmTrivia.Show
End Sub

Private Sub Command4_Click()
'Ends the program
End
End Sub

Private Sub Command5_Click()
'Takes the user to a home schedule viewer.
FrmMain.Hide
FrmSchedule.Show
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
