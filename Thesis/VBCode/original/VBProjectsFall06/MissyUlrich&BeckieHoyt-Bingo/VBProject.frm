VERSION 5.00
Begin VB.Form frmCreate 
   BackColor       =   &H00000000&
   Caption         =   "Create Your Own Bingo Board!"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   Picture         =   "VB Project.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNotWinner 
      Caption         =   "Didn't Win After 25 Numbers Have Been Called??  Click Here"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   41
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Click To Input Pre-Made Board"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   40
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtFree 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   39
      Text            =   "FREE"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtLastName 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   37
      Top             =   600
      Width           =   4575
   End
   Begin VB.CommandButton cmdWin 
      Caption         =   "        Have A Bingo?         Click Here"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   36
      Top             =   5520
      Width           =   3015
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   35
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   34
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox txtWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   1200
      Width           =   4935
   End
   Begin VB.TextBox txtO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   3960
      TabIndex        =   31
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3960
      TabIndex        =   30
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3960
      TabIndex        =   29
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3960
      TabIndex        =   28
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3960
      TabIndex        =   27
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3960
      TabIndex        =   26
      Text            =   "O"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   3120
      TabIndex        =   25
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3120
      TabIndex        =   24
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3120
      TabIndex        =   23
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3120
      TabIndex        =   22
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3120
      TabIndex        =   21
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3120
      TabIndex        =   20
      Text            =   "G"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtN 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   2280
      TabIndex        =   19
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtN 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2280
      TabIndex        =   18
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtN 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2280
      TabIndex        =   17
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtN 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2280
      TabIndex        =   16
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtN 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2280
      TabIndex        =   15
      Text            =   "N"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtI 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   1440
      TabIndex        =   14
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtI 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1440
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtI 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   1440
      TabIndex        =   12
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtI 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1440
      TabIndex        =   11
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtI 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtI 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1440
      TabIndex        =   9
      Text            =   "I"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   600
      TabIndex        =   8
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Text            =   "B"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdBoard 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create Your Own Board"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Bingo!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   5640
      TabIndex        =   1
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   5640
      TabIndex        =   0
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblLastName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please Enter Your Last Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please Enter Your First Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BingoProject
'Create Form
'Missy Ulrich & Beckie Hoyt
'November 3, 2006
'This form allows you to create your own bingo card by inputing numbers between certain values for each letter:
'B, I, N, G and O, or you can use a pre-made bingo card
'the user can input their name and switch between two forms to play the game
'The form also allows you to continue on, whether you have won or not at bingo, onto other forms to claim prizes
'This program allows you to play the game of bingo and receive prizes for winning or not
'This program is an interactive game of bingo because it allows the user to create their own scorecard, call the numbers, and finally check to see if they have a bingo or not
'If the user does or does not have a bingo after a certain amount of numbers have been called, the user has the option to continue on to the Winner or NotWinner forms of the program

Option Explicit

Private Sub cmdBoard_Click()
    Dim NewPos As Integer
    Dim C As Integer
    'Declare all the variables
    NumB = 1
    Do While NumB < 6
        TempB = InputBox("Please Enter A Number Between 1 and 15", "B Numbers")
        BArray(NumB) = TempB
            If TempB >= 1 And TempB <= 15 Then
                C = 0
                For NewPos = 1 To 5
                    If TempB = Val(txtB(NewPos).Text) Then
                        C = C + 1
                    End If
                Next NewPos
                If C = 0 Then
                    txtB(NumB).Text = "     " & TempB
                    NumB = NumB + 1
                Else
                    MsgBox ("You Have Already Entered That Number. Please Try Another Number")
                End If
            Else
                MsgBox ("Please Enter A Valid Number")
            End If
    Loop
    'This loop first receives input from the user through a inputbox and then checks to see if the number inputed by the user is between 1 and 15
    'The loop continues only with values 1-5 from the number of textboxes you can input items into on the form
    'If the number is between 1-15 and the number hasn't already been previously entered by the user, then the user can input another number.
    'This process of input continues 5 times until the B column on the card is full
    'If the number isn't between the values specified on the inputbox, then the user will be asked to enter another number
    NumI = 1
    Do While NumI < 6
        TempI = InputBox("Please Enter A Number Between 16 and 30", "I Numbers")
        IArray(NumI) = TempI
            If TempI >= 16 And TempI <= 30 Then
                C = 0
                For NewPos = 1 To 5
                    If TempI = Val(txtI(NewPos).Text) Then
                        C = C + 1
                    End If
                Next NewPos
                If C = 0 Then
                    txtI(NumI).Text = "     " & TempI
                    NumI = NumI + 1
                Else
                    MsgBox ("You Have Already Entered That Number. Please Try Another Number")
                End If
            Else
                MsgBox ("Please Enter A Valid Number")
            End If
    Loop
    'This loop first receives input from the user through a inputbox and then checks to see if the number inputed by the user is between 16 and 30
    'The loop continues only with values 1-5 from the number of textboxes you can input items into on the form
    'If the number is between 16-30 and the number hasn't already been previously entered by the user, then the user can input another number.
    'This process of input continues 5 times until the I column on the card is full
    'If the number isn't between the values specified on the inputbox, then the user will be asked to enter another number
    NumN = 1
    Do While NumN < 5
        TempN = InputBox("Please Enter A Number Between 31 and 45", "N Numbers")
        NArray(NumN) = TempN
            If TempN >= 31 And TempN <= 45 Then
                C = 0
                For NewPos = 1 To 4
                    If TempN = Val(txtN(NewPos).Text) Then
                        C = C + 1
                    End If
                Next NewPos
                If C = 0 Then
                    txtN(NumN).Text = "     " & TempN
                    NumN = NumN + 1
                Else
                    MsgBox ("You Have Already Entered That Number. Please Try Another Number")
                End If
            Else
                MsgBox ("Please Enter A Valid Number")
            End If
    Loop
    'This loop first receives input from the user through a inputbox and then checks to see if the number inputed by the user is between 31 and 45
    'The loop continues only with values 1-4 from the number of textboxes you can input items into on the form
    'If the number is between 31-45 and the number hasn't already been previously entered by the user, then the user can input another number.
    'This process of input continues 4 times until the N column on the card is full
    'If the number isn't between the values specified on the inputbox, then the user will be asked to enter another number
    NumG = 1
    Do While NumG < 6
        TempG = InputBox("Please Enter A Number Between 46 and 60", "G Numbers")
        GArray(NumG) = TempG
            If TempG >= 46 And TempG <= 60 Then
                C = 0
                For NewPos = 1 To 5
                    If TempG = Val(txtG(NewPos).Text) Then
                        C = C + 1
                    End If
                Next NewPos
                If C = 0 Then
                    txtG(NumG).Text = "     " & TempG
                    NumG = NumG + 1
                Else
                    MsgBox ("You Have Already Entered That Number. Please Try Another Number")
                End If
            Else
                MsgBox ("Please Enter A Valid Number")
            End If
    Loop
    'This loop first receives input from the user through a inputbox and then checks to see if the number inputed by the user is between 46 and 60
    'The loop continues only with values 1-5 from the number of textboxes you can input items into on the form
    'If the number is between 46-60 and the number hasn't already been previously entered by the user, then the user can input another number.
    'This process of input continues 5 times until the G column on the card is full
    'If the number isn't between the values specified on the inputbox, then the user will be asked to enter another number
    NumO = 1
    Do While NumO < 6
        TempO = InputBox("Please Enter A Number Between 61 and 75", "O Numbers")
        OArray(NumO) = TempO
            If TempO >= 61 And TempO <= 75 Then
                C = 0
                For NewPos = 1 To 5
                    If TempO = Val(txtO(NewPos).Text) Then
                        C = C + 1
                    End If
                Next NewPos
                If C = 0 Then
                    txtO(NumO).Text = "     " & TempO
                    NumO = NumO + 1
                Else
                    MsgBox ("You Have Already Entered That Number. Please Try Another Number")
                End If
            Else
                MsgBox ("Please Enter A Valid Number")
            End If
    Loop
    'This loop first receives input from the user through a inputbox and then checks to see if the number inputed by the user is between 61 and 75
    'The loop continues only with values 1-5 from the number of textboxes you can input items into on the form
    'If the number is between 61-75 and the number hasn't already been previously entered by the user, then the user can input another number.
    'This process of input continues 5 times until the O column on the card is full
    'If the number isn't between the values specified on the inputbox, then the user will be asked to enter another number
End Sub


Private Sub cmdExit_Click(Index As Integer)
    End
    'This allows the user to exit out of the program
End Sub

Private Sub cmdInput_Click()
    BNum = 1
    INum = 1
    NNum = 1
    GNum = 1
    ONum = 1
    NumB = 1
    NumI = 1
    NumN = 1
    NumG = 1
    NumO = 1
    'Sets the Variables to 1, which is the first textbox available for input
    Open App.Path & "\BingoNumbers.txt" For Input As #1
    Unit = 0
    Do Until EOF(1)
        Input #1, BNum, INum, NNum, GNum, ONum
        Unit = Unit + 1
        BArray(Unit) = BNum
        IArray(Unit) = INum
        NArray(Unit) = NNum
        GArray(Unit) = GNum
        OArray(Unit) = ONum
    Loop
    'This opens a text file from a different location, reads it and assigns variables to the data within the file
    'Stores the variables into the 5 different letter arrays : B, I, N, G and O
    Close #1
    Do While NumB < 6
        txtB(NumB).Text = "     " & BArray(NumB)
        NumB = NumB + 1
    Loop
    'The loop takes the data from the variables above that has been stored in the arrays and puts them into the correct locations of textboxes
    'The loop increments the textbox number to fill all the textboxes correctly, then stops when after the 5 boxes have been filled
    Do While NumI < 6
        txtI(NumI).Text = "     " & IArray(NumI)
        NumI = NumI + 1
    Loop
    'The loop takes the data from the variables above that has been stored in the arrays and puts them into the correct locations of textboxes
    'The loop increments the textbox number to fill all the textboxes correctly, then stops when after the 5 boxes have been filled
    Do While NumN < 5
        txtN(NumN).Text = "     " & NArray(NumN)
        NumN = NumN + 1
    Loop
    'The loop takes the data from the variables above that has been stored in the arrays and puts them into the correct locations of textboxes
    'The loop increments the textbox number to fill all the textboxes correctly, then stops when after the 4 boxes have been filled, because one
    'of the boxes is reserved for the free space
    Do While NumG < 6
        txtG(NumG).Text = "     " & GArray(NumG)
        NumG = NumG + 1
    Loop
    'The loop takes the data from the variables above that has been stored in the arrays and puts them into the correct locations of textboxes
    'The loop increments the textbox number to fill all the textboxes correctly, then stops when after the 5 boxes have been filled
    Do While NumO < 6
        txtO(NumO).Text = "     " & OArray(NumO)
        NumO = NumO + 1
    Loop
    'The loop takes the data from the variables above that has been stored in the arrays and puts them into the correct locations of textboxes
    'The loop increments the textbox number to fill all the textboxes correctly, then stops when after the 5 boxes have been filled
End Sub

Private Sub cmdNotWinner_Click()
    frmCreate.Hide
    frmNotWinner.Show
    'This switches from the Create form to the NotWinner form in the program
End Sub

Private Sub cmdPlay_Click(Index As Integer)
    frmCreate.Hide
    frmPlay.Show
    'This switches from the Create form to the Play form in the program
End Sub

Private Sub cmdStart_Click()
    UserName = txtName.Text
    UserLastName = txtLastName.Text
    User = Left(UserLastName, 1)
    txtWelcome.Text = ("Welcome to the game " & UserName & " " & User & "!")
    txtName.Visible = False
    txtLastName.Visible = False
    lblName.Visible = False
    lblLastName.Visible = False
    cmdStart.Visible = False
    'This is where the user enters their first and last name into two seperate textboxes
    'The program then takes the first name and the first letter of the last name and displays them both together in a textbox that Welcomes them to the game
    'The two textboxes where data was inputed by the user, and the two labels will also disappear
    
End Sub

Private Sub cmdWin_Click()
    frmCreate.Hide
    frmWinner.Show
    'This switches from the Create form to the Winner form in the program
End Sub


