VERSION 5.00
Begin VB.Form frmSavePrincess1 
   BackColor       =   &H00000000&
   Caption         =   "Save the Princess!"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   195
      Left            =   7440
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   735
      Left            =   7200
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Image img9 
      Height          =   2085
      Left            =   6720
      Picture         =   "frmSavePrincess1.frx":0000
      Top             =   3960
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Image img8 
      Height          =   2085
      Left            =   4560
      Picture         =   "frmSavePrincess1.frx":1287
      Top             =   3960
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Image img7 
      Height          =   2085
      Left            =   2400
      Picture         =   "frmSavePrincess1.frx":253C
      Top             =   3960
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Image img6 
      Height          =   2085
      Left            =   240
      Picture         =   "frmSavePrincess1.frx":3812
      Top             =   3960
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSavePrincess1.frx":4A98
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   8655
   End
   Begin VB.Image img5 
      Height          =   1215
      Index           =   3
      Left            =   -600
      Picture         =   "frmSavePrincess1.frx":4B34
      Top             =   5400
      Width           =   1380
   End
   Begin VB.Image img4 
      Height          =   1215
      Index           =   2
      Left            =   720
      Picture         =   "frmSavePrincess1.frx":5373
      Top             =   5400
      Width           =   1380
   End
   Begin VB.Image img3 
      Height          =   1215
      Index           =   1
      Left            =   2160
      Picture         =   "frmSavePrincess1.frx":5BB2
      Top             =   5400
      Width           =   1380
   End
   Begin VB.Image img2 
      Height          =   1215
      Index           =   0
      Left            =   3600
      Picture         =   "frmSavePrincess1.frx":63F1
      Top             =   5400
      Width           =   1380
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSavePrincess1.frx":6C30
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   3255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
   Begin VB.Image img1 
      Height          =   4620
      Left            =   4920
      Picture         =   "frmSavePrincess1.frx":6CF1
      Top             =   3000
      Width           =   2565
   End
End
Attribute VB_Name = "frmSavePrincess1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'This form allows the user to look into caves to get objects to
'put in the users inventory
'Save the princess1


Private Sub CmdNext_Click()
 lbl1.Visible = False
 img1.Visible = False
 img2(0).Visible = False
 img3(1).Visible = False
 img4(2).Visible = False
 img5(3).Visible = False
 cmdNext.Visible = False
 lbl2.Visible = True
 img6.Visible = True
 img7.Visible = True
 img8.Visible = True
 img9.Visible = True
 Inventory = ""
 
End Sub

Private Sub CmdQuit_Click()
    End
End Sub

Private Sub img6_Click()
 frmSavePrincess1.Hide
 frmCave1.Show
End Sub

Private Sub img7_Click()
 frmSavePrincess1.Hide
 frmCave2.Show
End Sub

Private Sub img8_Click()
 frmSavePrincess1.Hide
 frmCave3.Show
End Sub

Private Sub img9_Click()
 frmSavePrincess1.Hide
 frmCave4.Show
End Sub
