VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   Caption         =   "Wildlife Challenge"
   ClientHeight    =   8535
   ClientLeft      =   1365
   ClientTop       =   1095
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   12615
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00008000&
      Caption         =   "Take The Quiz"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00008000&
      Caption         =   "Minnesota Wildlife Information"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.PictureBox picMuleDeer 
      Height          =   4935
      Left            =   2760
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Label lblMe 
      BackColor       =   &H00008000&
      Caption         =   "Lance Uselman"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1695
      Left            =   2760
      Shape           =   2  'Oval
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label lblSportsmansChallenge 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Wildlife Challenge"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   6375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wildlife Challenge (Project1.vbp)
'frmMain (frmMain.frm)
'Lance Uselman
'March 24, 2006
'Purpose of the project: The purpose of this project is to quiz the user on
'                        different aspects of wildlife by asking a series of
'                        questions on a series of forms. The user can also
'                        access information on a couple of different species.
'Purpose of the form: This form allows the user to start the quiz, view
'                     information on a variety of species, or quit the program.

Option Explicit

Private Sub cmd1_Click()
    frmQ1.Show
    frmMain.Hide    'This button allows the user to access the quiz.
End Sub

Private Sub cmd2_Click()
    frmInfo.Show
    frmMain.Hide    'This button allows the user to view information on a variety of species.
End Sub

Private Sub cmd3_Click()
    End             'This button allows the user to exit the program.
End Sub

