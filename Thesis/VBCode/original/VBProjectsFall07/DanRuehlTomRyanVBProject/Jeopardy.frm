VERSION 5.00
Begin VB.Form frmJeopardy 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinalJeopardy 
      BackColor       =   &H00FF0000&
      Caption         =   "Go To Final Jeopardy"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2400
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6240
      Width           =   1455
   End
   Begin VB.PictureBox picresults1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   13080
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   38
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15120
      ScaleHeight     =   555
      ScaleWidth      =   2955
      TabIndex        =   36
      Top             =   9600
      Width           =   3015
   End
   Begin VB.CommandButton cmd6_1000 
      BackColor       =   &H00FF0000&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmd6_800 
      BackColor       =   &H00FF0000&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmd6_600 
      BackColor       =   &H00FF0000&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd6_400 
      BackColor       =   &H00FF0000&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmd6_200 
      BackColor       =   &H00FF0000&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmd5_1000 
      BackColor       =   &H00FF0000&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmd5_800 
      BackColor       =   &H00FF0000&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmd5_600 
      BackColor       =   &H00FF0000&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd5_400 
      BackColor       =   &H00FF0000&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmd5_200 
      BackColor       =   &H00FF0000&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmd4_1000 
      BackColor       =   &H00FF0000&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmd4_800 
      BackColor       =   &H00FF0000&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmd4_600 
      BackColor       =   &H00FF0000&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmd4_400 
      BackColor       =   &H00FF0000&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmd4_200 
      BackColor       =   &H00FF0000&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmd3_1000 
      BackColor       =   &H00FF0000&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmd3_800 
      BackColor       =   &H00FF0000&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmd3_600 
      BackColor       =   &H00FF0000&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd3_400 
      BackColor       =   &H00FF0000&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmd3_200 
      BackColor       =   &H00FF0000&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmd2_1000 
      BackColor       =   &H00FF0000&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmd_2800 
      BackColor       =   &H00FF0000&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmd_2600 
      BackColor       =   &H00FF0000&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd_2400 
      BackColor       =   &H00FF0000&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmd_2200 
      BackColor       =   &H00FF0000&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdOne1000 
      BackColor       =   &H00FF0000&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdOne800 
      BackColor       =   &H00FF0000&
      Caption         =   "800"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdOne600 
      BackColor       =   &H00FF0000&
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdOne400 
      BackColor       =   &H00FF0000&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdOne200 
      BackColor       =   &H00FF0000&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   13320
      TabIndex        =   37
      Top             =   7080
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   10110
      Left            =   12960
      Picture         =   "Jeopardy.frx":0000
      Top             =   120
      Width           =   5235
   End
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   0
      Picture         =   "Jeopardy.frx":E17C
      Top             =   7320
      Width           =   12885
   End
   Begin VB.Label lblCategory6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9720
      TabIndex        =   35
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblCategory5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7800
      TabIndex        =   34
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblCategory4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6000
      TabIndex        =   33
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblCategory3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4080
      TabIndex        =   32
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblCategory2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   31
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblCategory1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   30
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmJeopardy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_2200_Click()
        'Each of the buttons on this form perform the same few actions.
        'First, the button clears the score box so the updated score can be printed.
        'Then it makes the Question form appear and the main jeopardy form appear.
        'It then prints the question into the querstion label and the 4 answer options into thir respective buttons.
        'It prints the answer into an invisible text box
        'Finally, it makes the pressed button invisible.
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "This is the city where the White House is located."
    frmQuestion.cmdA.Caption = "A.  What is New York City."
    frmQuestion.cmdB.Caption = "B.  What is Philadelphia."
    frmQuestion.cmdC.Caption = "C.  What is Boston."
    frmQuestion.cmdD.Caption = "D.  What is Washington D.C."
    frmQuestion.txtAnswer = "D"
    F = 200
    cmd_2200.Visible = False
End Sub

Private Sub cmd_2400_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "The Taj Mahal is located in what Country."
    frmQuestion.cmdA.Caption = "A.  What is India."
    frmQuestion.cmdB.Caption = "B.  What is Iran."
    frmQuestion.cmdC.Caption = "C.  What is Iraq."
    frmQuestion.cmdD.Caption = "D.  What is France."
    frmQuestion.txtAnswer = "A"
    F = 400
    cmd_2400.Visible = False
End Sub

Private Sub cmd_2600_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "This is the tallest peak in Africa."
    frmQuestion.cmdA.Caption = "A.  What is Mt. McKinley."
    frmQuestion.cmdB.Caption = "B.  What is Mt. Everest."
    frmQuestion.cmdC.Caption = "C.  What is Mt. Kilimanjaro."
    frmQuestion.cmdD.Caption = "D.  What is Mt. Rushmore."
    frmQuestion.txtAnswer = "C"
    F = 600
    cmd_2600.Visible = False
End Sub

Private Sub cmd_2800_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "This city in Argentina is the southern most city in the world."
    frmQuestion.cmdA.Caption = "A.  What is Buenos Aires."
    frmQuestion.cmdB.Caption = "B.  What is Ushuaia."
    frmQuestion.cmdC.Caption = "C.  What is Brazilia"
    frmQuestion.cmdD.Caption = "D.  What is Satiago."
    frmQuestion.txtAnswer = "B"
    F = 800
    cmd_2800.Visible = False
End Sub

Private Sub cmd2_1000_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "This mountain is the tallest in the world measuring over 33,000 feet."
    frmQuestion.cmdA.Caption = "A.  What is Mauna Kea."
    frmQuestion.cmdB.Caption = "B.  What is Mt. Everest."
    frmQuestion.cmdC.Caption = "C.  What is Mt. Hood."
    frmQuestion.cmdD.Caption = "D.  What is K2."
    frmQuestion.txtAnswer = "A"
    F = 1000
    cmd2_1000.Visible = False
End Sub

Private Sub cmd3_1000_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "Light My Candle."
    frmQuestion.cmdA.Caption = "A.  What is Cats."
    frmQuestion.cmdB.Caption = "B.  What is Mama Mia."
    frmQuestion.cmdC.Caption = "C.  What is Rent."
    frmQuestion.cmdD.Caption = "D.  What is Phantom of the Opera."
    frmQuestion.txtAnswer = "C"
    F = 1000
    cmd3_1000.Visible = False
End Sub

Private Sub cmd3_200_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "Shapoopie."
    frmQuestion.cmdA.Caption = "A.  What is The Music Man."
    frmQuestion.cmdB.Caption = "B.  What is Oklahoma."
    frmQuestion.cmdC.Caption = "C.  What is High School Musical."
    frmQuestion.cmdD.Caption = "D.  What is Joesph and the Amazing Technicolor Dream Coat."
    frmQuestion.txtAnswer = "A"
    F = 200
    cmd3_200.Visible = False
End Sub

Private Sub cmd3_400_Click()
    
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "If I Were A Rich Man."
    frmQuestion.cmdA.Caption = "A.  What is Little Shop of Horrors."
    frmQuestion.cmdB.Caption = "B.  What is Guys and Dolls."
    frmQuestion.cmdC.Caption = "C.  What is Fiddler on the Roof."
    frmQuestion.cmdD.Caption = "D.  What is My Fair Lady."
    frmQuestion.txtAnswer = "C"
    F = 400
    cmd3_400.Visible = False
End Sub

Private Sub cmd3_600_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "Go Go Joseph."
    frmQuestion.cmdA.Caption = "A.  What Joseph and the Amazing Technicolor Dreamcoat."
    frmQuestion.cmdB.Caption = "B.  What is Oklahoma."
    frmQuestion.cmdC.Caption = "C.  What is High School Musical."
    frmQuestion.cmdD.Caption = "D.  What is Ty Cobb."
    frmQuestion.txtAnswer = "A"
    F = 600
    cmd3_600.Visible = False
End Sub

Private Sub cmd3_800_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "Getcha Head in the Game."
    frmQuestion.cmdA.Caption = "A.  What is High School Musical 2."
    frmQuestion.cmdB.Caption = "B.  What is Guys and Dolls."
    frmQuestion.cmdC.Caption = "C.  What is Grease."
    frmQuestion.cmdD.Caption = "D.  What is High School Musical."
    frmQuestion.txtAnswer = "D"
    F = 800
    cmd3_800.Visible = False
End Sub

Private Sub cmd4_1000_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "1977, Africa"
    frmQuestion.cmdA.Caption = "A.  Who is Todo."
    frmQuestion.cmdB.Caption = "B.  Who is Dexy's Midnight Runners."
    frmQuestion.cmdC.Caption = "C.  Who is Frankie Goes to Hollywood."
    frmQuestion.cmdD.Caption = "D.  Who is Thin Lizzy."
    frmQuestion.txtAnswer = "A"
    F = 1000
    cmd4_1000.Visible = False
End Sub

Private Sub cmd4_200_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "2002, Who Let the Dogs Out."
    frmQuestion.cmdA.Caption = "A.  Who are the Baha Men."
    frmQuestion.cmdB.Caption = "B.  Who is Blind Melon."
    frmQuestion.cmdC.Caption = "C.  Who is Billy Ray Cyrus."
    frmQuestion.cmdD.Caption = "D.  Who is House of Pain."
    frmQuestion.txtAnswer = "A"
    F = 200
    cmd4_200.Visible = False
End Sub

Private Sub cmd4_400_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "1997, Barbi Girl."
    frmQuestion.cmdA.Caption = "A.  Who is Ace of Base."
    frmQuestion.cmdB.Caption = "B.  Who is Tag Team."
    frmQuestion.cmdC.Caption = "C.  Who is Crash Test Dummies."
    frmQuestion.cmdD.Caption = "D.  Who is Aqua."
    frmQuestion.txtAnswer = "D"
    F = 400
    cmd4_400.Visible = False
End Sub

Private Sub cmd4_600_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "1998, C 'est la Vie"
    frmQuestion.cmdA.Caption = "A.  Who is Semisonic."
    frmQuestion.cmdB.Caption = "B.  Who is B*Witched."
    frmQuestion.cmdC.Caption = "C.  Who is Lou Bega."
    frmQuestion.cmdD.Caption = "D.  Who is Eifle 65."
    frmQuestion.txtAnswer = "B"
    F = 600
    cmd4_600.Visible = False
End Sub

Private Sub cmd4_800_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "1980, Turning Japanese"
    frmQuestion.cmdA.Caption = "A.  Who is Ah ha."
    frmQuestion.cmdB.Caption = "B.  Who is Devo."
    frmQuestion.cmdC.Caption = "C.  Who are The Vapors."
    frmQuestion.cmdD.Caption = "D.  Who is Taco."
    frmQuestion.txtAnswer = "C"
    F = 800
    cmd4_800.Visible = False
End Sub

Private Sub cmd5_1000_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "Chris Farley"
    frmQuestion.cmdA.Caption = "A.  What is USC."
    frmQuestion.cmdB.Caption = "B.  What is University of Oregon."
    frmQuestion.cmdC.Caption = "C.  What is Marquette University."
    frmQuestion.cmdD.Caption = "D.  What is LSU."
    frmQuestion.txtAnswer = "C"
    F = 1000
    cmd5_1000.Visible = False
End Sub

Private Sub cmd5_200_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "Bill Sexton"
    frmQuestion.cmdA.Caption = "A.  What is St. Thomas University"
    frmQuestion.cmdB.Caption = "B.  What is Harvard."
    frmQuestion.cmdC.Caption = "C.  What is University of Florida."
    frmQuestion.cmdD.Caption = "D.  What is St. John's University."
    frmQuestion.txtAnswer = "D"
    F = 200
    cmd5_200.Visible = False
End Sub

Private Sub cmd5_400_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "Michael Jordan"
    frmQuestion.cmdA.Caption = "A.  What is University of North Carolina."
    frmQuestion.cmdB.Caption = "B.  What is Duke University."
    frmQuestion.cmdC.Caption = "C.  What is Stanford University."
    frmQuestion.cmdD.Caption = "D.  What is UCLA."
    frmQuestion.txtAnswer = "A"
    F = 400
    cmd5_400.Visible = False
End Sub

Private Sub cmd5_600_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "George W. Bush"
    frmQuestion.cmdA.Caption = "A.  What is University of Texas Austin."
    frmQuestion.cmdB.Caption = "B.  What is Harvard."
    frmQuestion.cmdC.Caption = "C.  What is Never Graduated High School."
    frmQuestion.cmdD.Caption = "D.  What is Yale."
    frmQuestion.txtAnswer = "D"
    F = 600
    cmd5_600.Visible = False
End Sub

Private Sub cmd5_800_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "Mark Cuban"
    frmQuestion.cmdA.Caption = "A.  What is NYU."
    frmQuestion.cmdB.Caption = "B.  What is Indiana University."
    frmQuestion.cmdC.Caption = "C.  What is The University of Iowa."
    frmQuestion.cmdD.Caption = "D.  What is Baylor University."
    frmQuestion.txtAnswer = "B"
    F = 800
    cmd5_800.Visible = False
End Sub

Private Sub cmd6_1000_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "This company claims to be the first to can their beverages"
    frmQuestion.cmdA.Caption = "A.  What is Shasta."
    frmQuestion.cmdB.Caption = "B.  What is Pepsi."
    frmQuestion.cmdC.Caption = "C.  What is Coke."
    frmQuestion.cmdD.Caption = "D.  What is Mountain Dew."
    frmQuestion.txtAnswer = "A"
    F = 1000
    cmd6_1000.Visible = False
End Sub

Private Sub cmd6_200_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "This brand has christmas adds featuring Polar Bears."
    frmQuestion.cmdA.Caption = "A.  What is Pepsi."
    frmQuestion.cmdB.Caption = "B.  What is Mountain Dew."
    frmQuestion.cmdC.Caption = "C.  What is Coca Cola."
    frmQuestion.cmdD.Caption = "D.  What is Surge."
    frmQuestion.txtAnswer = "C"
    F = 200
    cmd6_200.Visible = False
End Sub

Private Sub cmd6_400_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "Lebron 'King' James sponsors this brand of pop."
    frmQuestion.cmdA.Caption = "A.  What is Mello Yellow."
    frmQuestion.cmdB.Caption = "B.  What is Sprite."
    frmQuestion.cmdC.Caption = "C.  What is Sierra Mist."
    frmQuestion.cmdD.Caption = "D.  What is Sun Drop."
    frmQuestion.txtAnswer = "B"
    F = 400
    cmd6_400.Visible = False
End Sub

Private Sub cmd6_600_Click()
    
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "This pop has the most caffeine in it."
    frmQuestion.cmdA.Caption = "A.  What is Sunkist."
    frmQuestion.cmdB.Caption = "B.  What is Pepsi."
    frmQuestion.cmdC.Caption = "C.  What is Mountain Dew."
    frmQuestion.cmdD.Caption = "D.  What is Jolt."
    frmQuestion.txtAnswer = "D"
    F = 600
    cmd6_600.Visible = False
End Sub

Private Sub cmd6_800_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "This major pop brand recently dropped Britteny Spears as a sponsor."
    frmQuestion.cmdA.Caption = "A.  What is Pepsi."
    frmQuestion.cmdB.Caption = "B.  What is Coke."
    frmQuestion.cmdC.Caption = "C.  What is Mountain Dew."
    frmQuestion.cmdD.Caption = "D.  What is Dr. Pepper."
    frmQuestion.txtAnswer = "A"
    F = 800
    cmd6_800.Visible = False
End Sub

Private Sub cmdFinalJeopardy_Click()
    'This button decides whether to take you to final jeopardy or not.
    'If the user's total is less than 1, they are taken to the quit form, their game is over.
    'If the user's total is 1 or more, they are asked how much they wish to wager.
    'If the user enters an invalid wager, an error meggage pops up and the user is asked to re-enter.
    'Finally, it takes you to the final jeopardy page.
    If Sum < 1 Then
        frmQuit.Show
        frmJeopardy.Hide
    Else
        MsgBox ("Let's go to final Jeopardy!")
        MsgBox ("Category is Ancient Tombs")
        Wager = InputBox("Enter the amount to wager, which must be less than or equal to your total amount.")
        frmFinal.Show
        frmJeopardy.Hide
        If Wager > Sum Or Wager < 0 Then
            MsgBox ("Invalid wager entry! Please enter an amount between 0 and your score.")
            Wager = InputBox("Enter the amount to wager.")
        End If
    End If
        
End Sub

Private Sub cmdOne1000_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "He is the youngest NHL player to win the scoring title."
    frmQuestion.cmdA.Caption = "A.  Who is Mario Lemuix."
    frmQuestion.cmdB.Caption = "B.  Who is Sydney Crosby."
    frmQuestion.cmdC.Caption = "C.  Who is Wayne Gretzky."
    frmQuestion.cmdD.Caption = "D.  Who is Alexander Ovechkin."
    frmQuestion.txtAnswer = "B"
    F = 1000
    cmdOne1000.Visible = False
End Sub

Private Sub cmdOne200_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide
   
    
    frmQuestion.lblQuestions.Caption = "He was the first African American baseball player."
    frmQuestion.cmdA.Caption = "A.  Who is Ted Williams."
    frmQuestion.cmdB.Caption = "B.  Who is Jackie Robinson."
    frmQuestion.cmdC.Caption = "C.  Who is Barry Bonds."
    frmQuestion.cmdD.Caption = "D.  Who is Ty Cobb."
    frmQuestion.txtAnswer = "B"
    F = 200
    
    cmdOne200.Visible = False
End Sub

Private Sub cmdOne400_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "He was the greatest player ever to play for the Chicago Bulls."
    frmQuestion.cmdA.Caption = "A.  Who is Steve Kerr."
    frmQuestion.cmdB.Caption = "B.  Who is Phil Jackson."
    frmQuestion.cmdC.Caption = "C.  Who is Michael Jordan."
    frmQuestion.cmdD.Caption = "D.  Who is Kevin Garnett."
    frmQuestion.txtAnswer = "C"
    F = 400
    cmdOne400.Visible = False
End Sub

Private Sub cmdOne600_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "What WNBA team has won the most titles."
    frmQuestion.cmdA.Caption = "A.  Who are the Houston Comets."
    frmQuestion.cmdB.Caption = "B.  Who are the Minnesota Lynx."
    frmQuestion.cmdC.Caption = "C.  Who are the Los Angeles Sparks."
    frmQuestion.cmdD.Caption = "D.  Who are the Detroit Shock."
    frmQuestion.txtAnswer = "A"
    F = 600
    cmdOne600.Visible = False
End Sub

Private Sub cmdOne800_Click()
    picResults.Cls
    frmQuestion.Show
    frmJeopardy.Hide

frmQuestion.lblQuestions.Caption = "The Center Court of the Australian Open in Tennis is named after."
    frmQuestion.cmdA.Caption = "A.  Who Andre Agassi."
    frmQuestion.cmdB.Caption = "B.  Who Leyton Hewitt."
    frmQuestion.cmdC.Caption = "C.  Who Bjorn Borg."
    frmQuestion.cmdD.Caption = "D.  Who Rod Laver."
    frmQuestion.txtAnswer = "D"
    F = 800
    cmdOne800.Visible = False
End Sub

Private Sub Command1_Click()
'This is the quit button, which simply ends the program.
End
End Sub


Private Sub Command2_Click()

End Sub
