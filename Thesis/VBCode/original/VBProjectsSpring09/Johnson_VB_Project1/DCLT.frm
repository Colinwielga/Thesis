VERSION 5.00
Begin VB.Form frmDCLT 
   BackColor       =   &H000000C0&
   Caption         =   "Drill Cadet Leadership Training"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   15555
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   360
      ScaleHeight     =   1455
      ScaleWidth      =   8055
      TabIndex        =   5
      Top             =   480
      Width           =   8055
   End
   Begin VB.CommandButton cmdWhy 
      BackColor       =   &H0000FFFF&
      Caption         =   "Why did I go here?"
      Height          =   855
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.PictureBox picDCLT1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   240
      ScaleHeight     =   4935
      ScaleWidth      =   8175
      TabIndex        =   3
      Top             =   2160
      Width           =   8175
   End
   Begin VB.PictureBox picDCLT2 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   9120
      ScaleHeight     =   3975
      ScaleWidth      =   6015
      TabIndex        =   2
      Top             =   1800
      Width           =   6015
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000FFFF&
      Caption         =   "Back to Main Form"
      Height          =   855
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit Program"
      Height          =   855
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Tim Johnson   3/20    To State Why I went to DCLT and to show a bit about it"
      Height          =   615
      Left            =   11880
      TabIndex        =   7
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Label lblDCLTTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "DCLT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10440
      TabIndex        =   6
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmDCLT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()     'Goes back to Main Form
frmMain.Show                    'Goes back to Main Form
frmDCLT.Hide
End Sub

Private Sub cmdQuit_Click()     'Ends program where you are
    End                         'Ends program where you are
End Sub

Private Sub cmdWhy_Click()      'Answers a simple question
picInfo.Print "I went to Drill Cadet Leadership Training after LDAC this past summer to, hopefully, get a better understanding"; Tab(3); " of what the basic course trainee and all enlisted soldiers go through."
picInfo.Print "    It was really an eye opening experience and I feel that I am better prepared for my future now that I have gone."

End Sub

Private Sub Form_Load()         'Puts ups pictures to improve form appearance

picDCLT1.Picture = LoadPicture(App.Path & "\" & dcltpix(2))
picDCLT2.Picture = LoadPicture(App.Path & "\" & dcltpix(1))

End Sub

