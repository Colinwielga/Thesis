VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H80000010&
   Caption         =   "PhotoMind™"
   ClientHeight    =   10485
   ClientLeft      =   570
   ClientTop       =   615
   ClientWidth     =   13875
   FillColor       =   &H00000080&
   ForeColor       =   &H80000001&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   13875
   Begin VB.TextBox txtName 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   11040
      TabIndex        =   10
      Top             =   9360
      Width           =   1815
   End
   Begin VB.CommandButton cmdStartNext 
      Caption         =   "Click here to test your Photo Mind"
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   9960
      Width           =   3975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000001&
      Caption         =   "Quit"
      Height          =   495
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   9960
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000010&
      Caption         =   "Please enter your name: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   11
      Top             =   9360
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   3945
      Left            =   8880
      Picture         =   "Form1_Intro.frx":0000
      Top             =   2160
      Width           =   3945
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000010&
      Caption         =   "Winner gets a prize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   8520
      Width           =   7455
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000010&
      Caption         =   "You can get help three times (the assist buttons at the bottom)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   8040
      Width           =   7455
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000010&
      Caption         =   "All answers will be scored at the end of the game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   7560
      Width           =   7455
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000010&
      Caption         =   "You will have 40 seconds to answer a question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   7080
      Width           =   7455
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000010&
      Caption         =   "You will be given a question with four answer choices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   6600
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000010&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000010&
      Caption         =   $"Form1_Intro.frx":12810
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2415
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13920
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000010&
      Caption         =   "Welcome to PhotoMind™"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   12255
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is an introductory form where the user is given quick description of the game and then is asked to write
'his name in a text box in order to proceed to the game. An error message will pop up if there is no name found
'in the text box. There is also a quit button. Ther latter buttons return all the game to the original state if
'the user decides to play it for the second time.

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdStartNext_Click()
'This button makes sure we got a name from user and saves it, if not it reminds to do so, then advances to the first question.
'If the user chooses to play the game again it resets the buttons to their original state
CTR = 0
Right = 0

'If/then to check if the user name has been entered
If txtName = "" Then
    MsgBox "Please enter your name to start the game", , "PhotoMind™"
Else
    
    PlayerName = txtName.Text
    frmQ1.Show
    frmIntro.Hide
End If

'resets variables to their original state, for user to be able to play again
frmQ1.cmdComputer.Enabled = True
frmQ2.cmdComputer.Enabled = True
frmQ3.cmdComputer.Enabled = True
frmQ4.cmdComputer.Enabled = True
frmQ5.cmdComputer.Enabled = True
frmQ6.cmdComputer.Enabled = True
Computer = 0

frmQ1.cmdFF.Enabled = True
frmQ2.cmdFF.Enabled = True
frmQ3.cmdFF.Enabled = True
frmQ4.cmdFF.Enabled = True
frmQ5.cmdFF.Enabled = True
frmQ6.cmdFF.Enabled = True
FF = 0

frmQ1.cmdGoogle.Enabled = True
frmQ2.cmdGoogle.Enabled = True
frmQ3.cmdGoogle.Enabled = True
frmQ4.cmdGoogle.Enabled = True
frmQ5.cmdGoogle.Enabled = True
frmQ6.cmdGoogle.Enabled = True

 
frmQ1.cmdA.Visible = True
frmQ2.cmdA.Visible = True
frmQ3.cmdA.Visible = True
frmQ4.cmdA.Visible = True
frmQ5.cmdA.Visible = True
frmQ6.cmdA.Visible = True

frmQ1.cmdB.Visible = True
frmQ2.cmdB.Visible = True
frmQ3.cmdB.Visible = True
frmQ4.cmdB.Visible = True
frmQ5.cmdB.Visible = True
frmQ6.cmdB.Visible = True

frmQ1.cmdC.Visible = True
frmQ2.cmdC.Visible = True
frmQ3.cmdC.Visible = True
frmQ4.cmdC.Visible = True
frmQ5.cmdC.Visible = True
frmQ6.cmdC.Visible = True

frmQ1.cmdD.Visible = True
frmQ2.cmdD.Visible = True
frmQ3.cmdD.Visible = True
frmQ4.cmdD.Visible = True
frmQ5.cmdD.Visible = True
frmQ6.cmdD.Visible = True


End Sub


