VERSION 5.00
Begin VB.Form frmFindInterests 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Best Fit Program by Interests"
   ClientHeight    =   9750
   ClientLeft      =   2535
   ClientTop       =   945
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   10770
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reset Quiz"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9000
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   600
      Picture         =   "frmFindInterests.frx":0000
      ScaleHeight     =   4935
      ScaleWidth      =   4095
      TabIndex        =   15
      Top             =   3720
      Width           =   4095
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Find Your Program(s)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "No "
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Fall"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Spring"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H00FFC0FF&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "More secluded"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmd8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nursing"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmd9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmd11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Yes please."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmd12 
      BackColor       =   &H00C0FFC0&
      Caption         =   "No Thanks."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmd14 
      BackColor       =   &H00FFC0C0&
      Caption         =   "centerally located area"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmd15 
      BackColor       =   &H00FFC0C0&
      Caption         =   "cabin type environment"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmd16 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmd17 
      BackColor       =   &H00FFFFC0&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find a Program That Best Fits Your Interests"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1815
      Left            =   600
      TabIndex        =   23
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblSeven 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7.  Would you like a foreign roommate?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   8040
      Width           =   5295
   End
   Begin VB.Label lblSix 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6.  Where would you prefer to live?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   21
      Top             =   6840
      Width           =   5295
   End
   Begin VB.Label lblFive 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5.  Would you like to learn a new language?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   5640
      Width           =   5295
   End
   Begin VB.Label lblFour 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4.  What is your major?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   19
      Top             =   4440
      Width           =   5295
   End
   Begin VB.Label lblThree 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3.  Which type of environment do you prefer?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   3240
      Width           =   5295
   End
   Begin VB.Label lblTwo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2.  Which semester do you prefer to be abroad?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   17
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Label lblOne 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1.  Have you taken a foreign language?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   840
      Width           =   5175
   End
End
Attribute VB_Name = "frmFindInterests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written 3/27/08 by Erika and Sammi

Option Explicit
Dim TotalA As Integer

'this form asks the user questions, with 2 command buttons for options as answers for each question
'each command button is worth either 1 or two points
'the total at the end of the quiz corresponds to certain study abroad programs best fits the user


Private Sub cmd1_Click()
TotalA = 0
cmd1.Visible = True
cmd2.Visible = False
TotalA = TotalA + 1
End Sub
Private Sub cmd2_Click()
cmd1.Visible = False
cmd2.Visible = True
TotalA = TotalA + 2
End Sub
Private Sub cmd3_Click()
cmd3.Visible = True
cmd4.Visible = False
TotalA = TotalA + 1
End Sub
Private Sub cmd4_Click()
cmd3.Visible = False
cmd4.Visible = True
TotalA = TotalA + 2
End Sub

Private Sub cmd6_Click()
cmd6.Visible = True
cmd7.Visible = False
TotalA = TotalA + 1
End Sub
Private Sub cmd7_Click()
cmd6.Visible = False
cmd7.Visible = True
TotalA = TotalA + 2
End Sub
Private Sub cmd8_Click()
cmd8.Visible = True
cmd9.Visible = False
TotalA = TotalA + 2
End Sub
Private Sub cmd9_Click()
cmd8.Visible = False
cmd9.Visible = True
TotalA = TotalA + 1
End Sub

Private Sub cmd11_Click()
cmd11.Visible = True
cmd12.Visible = False
TotalA = TotalA + 1
End Sub
Private Sub cmd12_Click()
cmd11.Visible = False
cmd12.Visible = True
TotalA = TotalA + 2
End Sub

Private Sub cmd14_Click()
cmd14.Visible = True
cmd15.Visible = False
TotalA = TotalA + 1
End Sub
Private Sub cmd15_Click()
cmd14.Visible = False
cmd15.Visible = True
TotalA = TotalA + 2
End Sub
Private Sub cmd16_Click()
cmd16.Visible = True
cmd17.Visible = False
TotalA = TotalA + 1
End Sub
Private Sub cmd17_Click()
cmd16.Visible = False
cmd17.Visible = True
TotalA = TotalA + 2
End Sub


Private Sub cmdFind_Click()
Select Case TotalA
 Case Is >= 14
    MsgBox ("You would probably best fit in the South Africa program! Enjoy the warm weather!")
 Case Is <= 7
    MsgBox ("You are a chameleon would get along in the France, Austria, Spain, Chile or the Greco-Roman programs! Feel free to check using our budget calculator for a more specific fit!")
 Case Else
    MsgBox ("You could study in Austria, Greece/Rome, Chile, London, Australia, China, Ireland, or Japan! Any of these would be fitting! Check your budget with our budget calculator to get a more specific fit!")
End Select

End Sub

Private Sub cmdReset_Click()
cmd1.Visible = True
cmd2.Visible = True
cmd3.Visible = True
cmd4.Visible = True
cmd6.Visible = True
cmd7.Visible = True
cmd8.Visible = True
cmd9.Visible = True
cmd11.Visible = True
cmd12.Visible = True
cmd14.Visible = True
cmd15.Visible = True
cmd16.Visible = True
cmd17.Visible = True

TotalA = 0

End Sub

Private Sub cmdGoBack_Click()
frmFindInterests.Hide
frmFind.Show
End Sub
