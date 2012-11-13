VERSION 5.00
Begin VB.Form frmVegas 
   BackColor       =   &H80000012&
   Caption         =   "Vegas"
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18255
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   18255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   3600
      ScaleHeight     =   855
      ScaleWidth      =   11895
      TabIndex        =   6
      Top             =   1080
      Width           =   11895
   End
   Begin VB.CommandButton cmdBlackJack 
      Caption         =   "Wanna play some Black Jack?"
      Height          =   735
      Left            =   1080
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdWhy 
      Caption         =   "Why Did I Go?"
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.PictureBox picVegas 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   3600
      ScaleHeight     =   7935
      ScaleWidth      =   11895
      TabIndex        =   2
      Top             =   1320
      Width           =   11895
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Form"
      Height          =   855
      Left            =   10800
      TabIndex        =   1
      Top             =   9600
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit Program"
      Height          =   855
      Left            =   12840
      TabIndex        =   0
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Label lblShoutOut 
      BackColor       =   &H0000FFFF&
      Caption         =   "OF COURSE YOU WANNA PLAY SOME BLACK JACK"
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
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label lblObjective 
      BackColor       =   &H80000012&
      Caption         =   "Tim Johnson    3/20     To show why i went to vegas, give a little info about it, and to let the user get to the Black Jack game"
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   15480
      TabIndex        =   7
      Top             =   9600
      Width           =   2535
   End
   Begin VB.Label lblVegasTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Las Vegas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   7920
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmVegas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()         'Goes back to Main Form
frmMain.Show                        'Goes back to Main Form
frmVegas.Hide
End Sub

Private Sub cmdBlackJack_Click()    'Goes back to Black Jack Form
frmBlackJack.Show                   'Goes back to Black Jack Form
frmVegas.Hide
End Sub

Private Sub cmdQuit_Click()         'Ends program where you are
    End                             'Ends program where you are
End Sub

Private Sub cmdWhy_Click()          'Answers a simple question

picInfo.Print "I went because my best friend and I turned 21 in June, 2007, and his aunt was able to get us a great deal at the Jocky Club through her time share: only $50 a piece for the whole week!!"

End Sub

Private Sub Form_Load()             'Puts up a picturea to improve form appearance

picVegas.Picture = LoadPicture(App.Path & "\" & vegaspix(1))

End Sub
