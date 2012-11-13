VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H000040C0&
   Caption         =   "Main Page"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   4  'Icon
   ScaleHeight     =   7320
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuest 
      Caption         =   "Guest Book"
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
      Left            =   360
      TabIndex        =   5
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      TabIndex        =   4
      Top             =   6360
      Width           =   1215
   End
   Begin VB.PictureBox picMain 
      Height          =   5055
      Left            =   2760
      Picture         =   "Form1.frx":247D2
      ScaleHeight     =   4995
      ScaleWidth      =   3795
      TabIndex        =   3
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdfrmPic 
      Caption         =   "Artist's Information"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdFrmArt 
      Caption         =   "Artwork"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      MaskColor       =   &H0080FFFF&
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Art Portfolio"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Ashley Thompson"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   4815
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

Private Sub cmdGuest_Click()
frmMain.Hide
frmGuest.Show
End Sub

Private Sub cmdQuit1_Click()
End
End Sub

