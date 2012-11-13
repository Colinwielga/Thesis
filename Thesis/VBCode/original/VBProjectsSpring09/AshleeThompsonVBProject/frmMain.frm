VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H000040C0&
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      TabIndex        =   6
      Top             =   7440
      Width           =   1215
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   3480
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   4905
      ScaleWidth      =   3825
      TabIndex        =   5
      Top             =   2760
      Width           =   3855
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   4
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdfrmPic 
      Caption         =   "Artist's Information"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdFrmArt 
      Caption         =   "Artwork"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Art Portfolio"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Ashley Thompson"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Artist's Multimedia Portfolio
'frmMain
'Ashley Thompson
'Friday March 20, 2009
'This form acts as a Main Menu for the project. It allows the user to navigate to different areas of the program
'using the .Show and .Hide function
'It also has a Quit button


Private Sub cmdConnect_Click()
frmMain.Hide
frmConnect.Show
End Sub

Private Sub cmdFrmArt_Click()
frmMain.Hide
frmArtMain.Show
End Sub


Private Sub cmdfrmPic_Click()
frmMain.Hide
frmPic.Show
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub cmdQuit1_Click()
End
End Sub
