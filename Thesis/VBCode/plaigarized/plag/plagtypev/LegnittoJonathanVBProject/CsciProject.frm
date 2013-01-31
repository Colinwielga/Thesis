VERSION 5.00
Begin VB.Form FormHome
   BackColor       =   &H00C00000&
   Caption         =   "Toy Story"
   ClientHeight    =   12900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20550
   FillColor       =   &H00FF0000&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12900
   ScaleWidth      =   20550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPin
      BackColor       =   &H0000FF00&
      Caption         =   "Pinocchio"
      BeginProperty Font
         Name            =   "Mathematica6"
         Size            =   36
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9480
      Width           =   3735
   End
   Begin VB.PictureBox Picture2
      Height          =   4455
      Left            =   1200
      Picture         =   "CsciProject.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   4515
      TabIndex        =   7
      Top             =   4680
      Width           =   4575
   End
   Begin VB.PictureBox Picture1
      Height          =   4695
      Left            =   7800
      Picture         =   "CsciProject.frx":402D2
      ScaleHeight     =   4635
      ScaleWidth      =   5235
      TabIndex        =   5
      Top             =   4680
      Width           =   5295
   End
   Begin VB.CommandButton cmdMemoryGame
      BackColor       =   &H0000FF00&
      Caption         =   "Aladdin"
      BeginProperty Font
         Name            =   "Script MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9480
      Width           =   3015
   End
   Begin VB.CommandButton cmdToyStory
      BackColor       =   &H0000FF00&
      Caption         =   "Toy Story "
      BeginProperty Font
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15840
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   3255
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Exit"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   17760
      Picture         =   "CsciProject.frx":DA57C
      TabIndex        =   0
      Top             =   11880
      Width           =   2655
   End
   Begin VB.Label Label5
      BackColor       =   &H00C00000&
      Caption         =   "Identify your Toy Story Friends!"
      BeginProperty Font
         Name            =   "MS UI Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   14760
      TabIndex        =   10
      Top             =   4560
      Width           =   5535
   End
   Begin VB.Label Label4
      BackColor       =   &H00C00000&
      Caption         =   "Measure the lies of Pinocchio!"
      BeginProperty Font
         Name            =   "MS UI Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   4080
      Width           =   5055
   End
   Begin VB.Label Label3
      BackColor       =   &H00C00000&
      Caption         =   "Play Concentration With Aladdin!"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   7680
      TabIndex        =   6
      Top             =   4200
      Width           =   6135
   End
   Begin VB.Label Label2
      BackColor       =   &H00C00000&
      Caption         =   "Choose a movie to play a game with the characters from that movie"
      BeginProperty Font
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   16215
   End
   Begin VB.Label Label1
      BackColor       =   &H00C00000&
      Caption         =   "Disney Games"
      BeginProperty Font
         Name            =   "Lucida Handwriting"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   14775
   End
   Begin VB.Shape FrameToyStory
      BorderColor     =   &H8000000D&
      BorderWidth     =   25
      Height          =   3135
      Left            =   15840
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Image Image1
      Height          =   2790
      Left            =   15840
      Picture         =   "CsciProject.frx":111D74
      Top             =   5520
      Width           =   3180
   End
End
Attribute VB_Name = "FormHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Disney Games
'Form Name: FormHome
'Author: Jonathan Legnitto
'2/25/10
'Objective: The objective of this program was to make simple games that little kids who like disney could play that had complex code behind them
            'I put a bunch of images throughout the program to entertain kids and I filled all the forms with lots of color
Option Explicit
Dim Found As Boolean

Private Sub cmdPin_Click()
If Found = False Then                                       'This statement makes it so that the user is only asked for their name once
    MsgBox ("Whoa whoa whoa...not so fast! What's Your Name?")

    Username = InputBox("Please enter your name")

    MsgBox ("Hey there " & Username & "! Welcome to the wonderful world of Disney Games! Have Fun!")

    Found = True
End If
FormHome.Hide
FormPin.Show
FormMemory.Hide
FormToyStory.Hide


End Sub

Private Sub cmdQuit_Click()
End
End Sub
Found = False
Private Sub cmdMemoryGame_Click()
If Found = False Then                                       'This statement makes it so that the user is only asked for their name once
    MsgBox ("Whoa whoa whoa...not so fast! What's Your Name?")

    Username = InputBox("Please enter your name")

    MsgBox ("Hey there " & Username & "! Welcome to the wonderful world of Disney Games! Have Fun!")

    Found = True
End If
FormMemory.Show
FormHome.Hide
FormToyStory.Hide
FormPin.Hide

MsgBox (Username & ", Match the Images, Click Start/Reset to begin.")
End Sub



Private Sub cmdToyStory_Click()
If Found = False Then                                       'This statement makes it so that the user is only asked for their name once
    MsgBox ("Whoa whoa whoa...not so fast! What's Your Name?")

    Username = InputBox("Please enter your name")

    MsgBox ("Hey there " & Username & "! Welcome to the wonderful world of Disney Games! Have Fun!")

    Found = True
End If
FormMemory.Hide
FormHome.Hide
FormToyStory.Show
FormPin.Hide
Do Until True = True
Loop
Do Until True = True
Loop
Do Until True = True
Loop
Do Until True = True
Loop
End Sub

