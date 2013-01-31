VERSION 5.00
Begin VB.Form frmPoll
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1
      Height          =   3255
      Left            =   7080
      Picture         =   "frmPoll.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   2955
      TabIndex        =   7
      Top             =   2640
      Width           =   3015
   End
   Begin VB.CommandButton cmdHome
      BackColor       =   &H0000C000&
      Caption         =   "Go to Main Menu"
      BeginProperty Font
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.PictureBox etrnhetr
      BeginProperty Font
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      ScaleHeight     =   1035
      ScaleWidth      =   6270
      TabIndex        =   5
      Top             =   1080
      Width           =   6330
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H0000C000&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdNext
      BackColor       =   &H0000C000&
      Caption         =   "Go to Next Form"
      BeginProperty Font
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCompute
      BackColor       =   &H0000C000&
      Caption         =   "Compute Awesomeness"
      BeginProperty Font
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtInput
      Alignment       =   2  'Center
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "On a scale of 1 to 100, how much do you love the Packers?"
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Shape Shape1
      FillColor       =   &H00008000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   6495
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmPoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Get to know the Packers' Receivers
'frmPoll
'Sam Pederson
'2/23/10
'This form tells you how awesome you are

Private Sub cmdCompute_Click() 'this button computes your awesomeness
    Dim hhhdd As Double
    hhhdd = txtInput.Text
    etrnhetr.Cls
    Select Case hhhdd
        Case Is >= 90
            etrnhetr.Print "Good Choice. Go Pack Go!"
        Case 70 To 89
            etrnhetr.Print "You can do a little better than that!"
        Case 40 To 69
            etrnhetr.Print "You must not know football very well."
        Case 0 To 39
            etrnhetr.Print "Well hey there Viking fan, how are those Super Bowls treating ya?"
        Case Else
            etrnhetr.Print "What are you doing?"
    End Select
End Sub

Private Sub cmdHome_Click() 'this button takes you to the menu
    frmWelcome.Hide
    frmMenu.Show
    qqq.Hide
    frmData.Hide
    ttt.Hide
    frmPics.Hide
    frmMusic.Hide
    frmLast.Hide
End Sub

Private Sub cmdNext_Click() 'this button takes you to the next form
    frmWelcome.Hide
    frmMenu.Hide
    qqq.Hide
    frmData.Show
    ttt.Hide
    frmPics.Hide
    frmMusic.Hide
    frmLast.Hide
End Sub

Private Sub eeer_Click() 'this button ends the program
    End
End Sub
