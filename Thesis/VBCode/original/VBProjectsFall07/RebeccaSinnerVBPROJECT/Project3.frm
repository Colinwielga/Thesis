VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form3"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   8625
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Back to Selection Screen"
      Height          =   855
      Left            =   7920
      TabIndex        =   10
      Top             =   7560
      Width           =   2535
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   7920
      Width           =   1815
   End
   Begin VB.TextBox txt123 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4200
      TabIndex        =   4
      Top             =   7440
      Width           =   2655
   End
   Begin VB.PictureBox pix123 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   2640
      Picture         =   "Project3.frx":0000
      ScaleHeight     =   4095
      ScaleWidth      =   5775
      TabIndex        =   3
      Top             =   0
      Width           =   5775
   End
   Begin VB.PictureBox pixbed 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      Picture         =   "Project3.frx":141B6
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.PictureBox pixcar 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4560
      Picture         =   "Project3.frx":14E17
      ScaleHeight     =   1095
      ScaleWidth      =   1935
      TabIndex        =   1
      Top             =   4800
      Width           =   1935
   End
   Begin VB.PictureBox pixsweatshirt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8520
      Picture         =   "Project3.frx":15B5C
      ScaleHeight     =   1815
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblCar 
      BackColor       =   &H0000FFFF&
      Caption         =   "  2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   8
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblsweatshirt 
      BackColor       =   &H0000FFFF&
      Caption         =   "  3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblBed 
      BackColor       =   &H0000FFFF&
      Caption         =   "  1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblentry 
      BackColor       =   &H0000FFFF&
      Caption         =   $"Project3.frx":170BA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   6480
      Width           =   7695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit             'In this game, the user is shown three prizes which are numbered 1-3.
Dim Entry As Double         'They must type in the prizes in order from least to most expensive using each product's corresponding number.
                            'If they are correct, they "win" all three prizes.  If they are incorrect, the game is over and they can play again, go back to the selection screen, or quit playing.
                            'Under Option Explicit, I declared my variables.

Private Sub cmdEnter_Click()    'The user presses this button once they have made their guess to see if they are correct.
Entry = txt123                  'The user enters their guess in this text box.
Select Case Entry
    Case Is = 312
        MsgBox "Congratulations!  You win all three prizes!"
        Form2.Show  'Takes the player back to the selection screen if they win.
        Form3.Hide
    Case Is = 321
        MsgBox "I'm sorry, the correct order was 3-1-2.  You lose!"
    Case Is = 123
        MsgBox "I'm sorry, the correct order was 3-1-2.  You lose!"
    Case Is = 132
        MsgBox "I'm sorry, the correct order was 3-1-2.  You lose!"
    Case Is = 213
        MsgBox "I'm sorry, the correct order was 3-1-2.  You lose!"
    Case Is = 231
        MsgBox "I'm sorry, the correct order was 3-1-2.  You lose!"
    Case Else
        MsgBox "Invalid Entry. Remember: enter only the numbers 1,2,and 3 in their correct order with NO spaces, commas, etc."
    End Select
End Sub


Private Sub cmdQuit_Click()  'This button us used so the player can go back to Form2 and choose a different game or quit the game if they wish.
Form2.Show
Form3.Hide
End Sub
