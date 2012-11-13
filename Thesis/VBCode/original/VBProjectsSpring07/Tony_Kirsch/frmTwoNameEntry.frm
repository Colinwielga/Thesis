VERSION 5.00
Begin VB.Form frmTwoNameEntry 
   BackColor       =   &H00000000&
   Caption         =   "Please Enter Your Name"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   11085
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdphasethree 
      BackColor       =   &H0000FF00&
      Caption         =   "Continue on"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9360
      Width           =   3015
   End
   Begin VB.CommandButton cmdentername 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter your name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8280
      Width           =   3015
   End
   Begin VB.PictureBox picBTK 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   6360
      Picture         =   "frmTwoNameEntry.frx":0000
      ScaleHeight     =   2295
      ScaleWidth      =   1815
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picGacy 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1200
      Picture         =   "frmTwoNameEntry.frx":249C
      ScaleHeight     =   2295
      ScaleWidth      =   2535
      TabIndex        =   3
      Top             =   0
      Width           =   2535
   End
   Begin VB.PictureBox picFish 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   11520
      Picture         =   "frmTwoNameEntry.frx":43D8
      ScaleHeight     =   2415
      ScaleWidth      =   2175
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox picDahmer 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   8400
      Picture         =   "frmTwoNameEntry.frx":658A
      ScaleHeight     =   2055
      ScaleWidth      =   2655
      TabIndex        =   1
      Top             =   3960
      Width           =   2655
   End
   Begin VB.PictureBox picBundy 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   3720
      Picture         =   "frmTwoNameEntry.frx":83C6
      ScaleHeight     =   2055
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label lblDahmer 
      BackColor       =   &H00000000&
      Caption         =   $"frmTwoNameEntry.frx":950E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   7440
      TabIndex        =   9
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label lblbundy 
      BackColor       =   &H00000000&
      Caption         =   $"frmTwoNameEntry.frx":95B2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   2400
      TabIndex        =   8
      Top             =   6240
      Width           =   4215
   End
   Begin VB.Label lblFish 
      BackColor       =   &H00000000&
      Caption         =   $"frmTwoNameEntry.frx":9656
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   10200
      TabIndex        =   7
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label lblBTK 
      BackColor       =   &H00000000&
      Caption         =   "Dennis Lynn Rader ""BTK"": Murdered at least 10 people"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   6000
      TabIndex        =   6
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblGacy 
      BackColor       =   &H00000000&
      Caption         =   "John Wayne Gacy: Raped and murdered 33 boys and young males."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   3015
   End
End
Attribute VB_Name = "frmTwoNameEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form contains pictures of 5 of the most famous killlers and
'rapists and pedophiles that have ever lived. I took the images from the Crime Libray
'database and the little facts about them. Also the user will be told to input a name
'and this is will be used throughout the program



Private Sub cmdentername_Click()
'nam is declared globally so i just made name equal to whatever the user types
'in in the input box.
    nam = InputBox("Type in what you want to be called please.", "Enter Name")
    
End Sub

Private Sub cmdphasethree_Click()
'Takes the user to the case file form in which they can start their profiling
    frmTwoNameEntry.Hide
    frmCasefiles.Show
End Sub

