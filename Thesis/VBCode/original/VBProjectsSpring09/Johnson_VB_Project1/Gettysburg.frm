VERSION 5.00
Begin VB.Form frmGettysburg 
   BackColor       =   &H8000000D&
   Caption         =   "Gettysburg"
   ClientHeight    =   11625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17355
   LinkTopic       =   "Form1"
   ScaleHeight     =   11625
   ScaleWidth      =   17355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   5520
      ScaleHeight     =   3375
      ScaleWidth      =   3855
      TabIndex        =   7
      Top             =   7080
      Width           =   3855
   End
   Begin VB.CommandButton cmdWhy 
      Caption         =   "Why I Went Here?"
      Height          =   855
      Left            =   7080
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.PictureBox picGettysburg3 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   360
      ScaleHeight     =   7095
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   3360
      Width           =   4815
   End
   Begin VB.PictureBox picGettysburg2 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   9720
      ScaleHeight     =   4935
      ScaleWidth      =   6255
      TabIndex        =   3
      Top             =   5520
      Width           =   6255
   End
   Begin VB.PictureBox picGettysburg1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   5640
      ScaleHeight     =   4815
      ScaleWidth      =   9015
      TabIndex        =   2
      Top             =   480
      Width           =   9015
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Form"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit Program"
      Height          =   855
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Tim Johnson   3/20    To State Why I went to Gettysburg and to show a bit about it"
      Height          =   615
      Left            =   13680
      TabIndex        =   8
      Top             =   10680
      Width           =   2895
   End
   Begin VB.Label lblGettysburgTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Gettysburg, Pennsylvania"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   960
      TabIndex        =   5
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "frmGettysburg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()       'Goes back to Main Form
frmMain.Show                      'Goes back to Main Form
frmGettysburg.Hide
End Sub

Private Sub cmdQuit_Click()      'Ends program where you are
    End                          'Ends program where you are
End Sub

Private Sub cmdWhy_Click()       'Answers a simple question

picInfo.Print "I went to Gettysburg with my SJU ROTC Senior"
picInfo.Print "  classmates and teacher. We Spent several months"
picInfo.Print "  researching the overall battle and individual"
picInfo.Print "  leaders in the battle. The entire process is"
picInfo.Print "  called a battle analysis and is part of a staff "
picInfo.Print "  ride that every senior ROTC cadet is supposed"
picInfo.Print "  to do. Staff rides and battlefield analysises"
picInfo.Print "  are commonplace in the Army."

End Sub

Private Sub Form_Load()           'Puts ups pictures to improve form appearance

picGettysburg1.Picture = LoadPicture(App.Path & "\" & gettysburgpix(1))
picGettysburg2.Picture = LoadPicture(App.Path & "\" & gettysburgpix(2))
picGettysburg3.Picture = LoadPicture(App.Path & "\" & gettysburgpix(3))

End Sub
