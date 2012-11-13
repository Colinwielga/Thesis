VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form4"
   Picture         =   "Journey.frx":0000
   ScaleHeight     =   9360
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Screw this! Let's go Hunting!"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1440
      TabIndex        =   32
      Top             =   5520
      Width           =   3135
   End
   Begin VB.CommandButton cmdA1 
      Caption         =   "Grueling"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdB1 
      Caption         =   "Quick"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   23
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdC1 
      Caption         =   "Steady"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdD1 
      Caption         =   "Slow"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   2775
      Left            =   6000
      ScaleHeight     =   2715
      ScaleWidth      =   7155
      TabIndex        =   20
      Top             =   600
      Width           =   7215
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Click to See Your FATE"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   720
      TabIndex        =   19
      Top             =   7680
      Width           =   4335
   End
   Begin VB.PictureBox picResults2 
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   6000
      ScaleHeight     =   4275
      ScaleWidth      =   7155
      TabIndex        =   18
      Top             =   3480
      Width           =   7215
   End
   Begin VB.CommandButton cmdA2 
      Caption         =   "Hearty"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4320
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdB2 
      Caption         =   "Meager"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdC2 
      Caption         =   "Sparing"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdD2 
      Caption         =   "Satisfying"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3000
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdA3 
      Caption         =   "Wait "
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4320
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdB3 
      Caption         =   "Ferry"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdC3 
      Caption         =   "Ford"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdD3 
      Caption         =   "Pay Indian"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdA4 
      Caption         =   "Rest"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdB4 
      Caption         =   "Eat Herbs"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdC4 
      Caption         =   "Quicken Pace"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4320
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdD4 
      Caption         =   "Endure"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   6
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdA5 
      Caption         =   "Raft!"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdB5 
      Caption         =   "Winter Over"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1560
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdC5 
      Caption         =   "Continue by Land"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3120
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdD5 
      Caption         =   "Take a Taxi"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Home"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   1
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit the Game (cause you need to)"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   0
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the Trail"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label lblQ1 
      BackStyle       =   0  'Transparent
      Caption         =   "1.  Choose Your Pace!"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label lblQ2 
      BackStyle       =   0  'Transparent
      Caption         =   "2.  Choose Your Meal Portions!"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   2280
      Width           =   5415
   End
   Begin VB.Label lblQ3 
      BackStyle       =   0  'Transparent
      Caption         =   "3. How Will You Cross the River?"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Label lblQ4 
      BackStyle       =   0  'Transparent
      Caption         =   "4.  What Will You Do When You Get Sick?"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   3720
      Width           =   5295
   End
   Begin VB.Label lblQ5 
      BackStyle       =   0  'Transparent
      Caption         =   "5.  How Will You End Your Journey?"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   4440
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click the option you deem to posses the Best Strategery"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   25
      Top             =   720
      Width           =   5895
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Left            =   6000
      Top             =   600
      Width           =   6615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   4335
      Left            =   6000
      Top             =   3480
      Width           =   4215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Drew Madson & Sam Erickson Oct 2008
'This progam is a modified version of the Harry Potter sorting hat
'portion of
'N:\Classes\CS130\Trutwin_VB_Examples\Project Stuff\Sample Projects\A Virtual Hogwarts
'this is the game part of Oregon Trail. It uses buttons to tally all possible outcomes

Dim TotalA As Integer 'sets/stores value of TotalA as an integer
Dim TotalB As Integer 'sets/stores value of TotalB as an integer
Dim TotalC As Integer 'sets/stores value of TotalC as an integer
Dim TotalD As Integer 'sets/stores value of TotalD as an integer

Private Sub cmdA1_Click()
    TotalA = 0 'gives TotalA an initial value of 0
    TotalB = 0 'sets initial value of TotalB to 0
    TotalC = 0 'sets initial value of TotalC to 0
    TotalD = 0 'sets initial value of TotalD to 0
    cmdA1.Visible = True
    cmdB1.Visible = False
    cmdC1.Visible = False
    cmdD1.Visible = False
    TotalA = TotalA + 1 'adds 1 to TotalA
End Sub

Private Sub cmdA2_Click()
    cmdA2.Visible = True
    cmdB2.Visible = False
    cmdC2.Visible = False
    cmdD2.Visible = False
    TotalA = TotalA + 1 'adds 1 to TotalA
End Sub

Private Sub cmdA3_Click()
    cmdA3.Visible = True
    cmdB3.Visible = False
    cmdC3.Visible = False
    cmdD3.Visible = False
    TotalA = TotalA + 1 'adds 1 to TotalA
End Sub

Private Sub cmdA4_Click()
    cmdA4.Visible = True
    cmdB4.Visible = False
    cmdC4.Visible = False
    cmdD4.Visible = False
    TotalA = TotalA + 1 'adds 1 to TotalA
End Sub

Private Sub cmdA5_Click()
    cmdA5.Visible = True
    cmdB5.Visible = False
    cmdC5.Visible = False
    cmdD5.Visible = False
    TotalA = TotalA + 1 'adds 1 to TotalA
End Sub

Private Sub cmdB1_Click()
    TotalA = 0 'gives TotalA an initial value of 0
    TotalB = 0 'sets initial value of TotalB to 0
    TotalC = 0 'sets initial value of TotalC to 0
    TotalD = 0 'sets initial value of TotalD to 0
    cmdB1.Visible = True
    cmdA1.Visible = False
    cmdC1.Visible = False
    cmdD1.Visible = False
    TotalB = TotalB + 1 'adds 1 to TotalB
End Sub

Private Sub cmdB2_Click()
    cmdA2.Visible = False
    cmdB2.Visible = True
    cmdC2.Visible = False
    cmdD2.Visible = False
    TotalB = TotalB + 1 'adds 1 to TotalB
End Sub

Private Sub cmdB3_Click()
    cmdA3.Visible = False
    cmdB3.Visible = True
    cmdC3.Visible = False
    cmdD3.Visible = False
    TotalB = TotalB + 1 'adds 1 to TotalB
End Sub

Private Sub cmdB4_Click()
    cmdA4.Visible = False
    cmdB4.Visible = True
    cmdC4.Visible = False
    cmdD4.Visible = False
    TotalB = TotalB + 1 'adds 1 to TotalB
End Sub

Private Sub cmdB5_Click()
    cmdA5.Visible = False
    cmdB5.Visible = True
    cmdC5.Visible = False
    cmdD5.Visible = False
    TotalB = TotalB + 1 'adds 1 to TotalB
End Sub

Private Sub cmdC1_Click()
    TotalA = 0 'gives TotalA an initial value of 0
    TotalB = 0 'sets initial value of TotalB to 0
    TotalC = 0 'sets initial value of TotalC to 0
    TotalD = 0 'sets initial value of TotalD to 0
    cmdC1.Visible = True
    cmdB1.Visible = False
    cmdA1.Visible = False
    cmdD1.Visible = False
    TotalC = TotalC + 1 'adds 1 to TotalC
End Sub

Private Sub cmdC2_Click()
    cmdA2.Visible = False
    cmdB2.Visible = False
    cmdC2.Visible = True
    cmdD2.Visible = False
    TotalC = TotalC + 1 'adds 1 to TotalC
End Sub

Private Sub cmdC3_Click()
    cmdA3.Visible = False
    cmdB3.Visible = False
    cmdC3.Visible = True
    cmdD3.Visible = False
    TotalC = TotalC + 1 'adds 1 to TotalC
End Sub

Private Sub cmdC4_Click()
    cmdA4.Visible = False
    cmdB4.Visible = False
    cmdC4.Visible = True
    cmdD4.Visible = False
    TotalC = TotalC + 1 'adds 1 to TotalC
End Sub

Private Sub cmdC5_Click()
    cmdA5.Visible = False
    cmdB5.Visible = False
    cmdC5.Visible = True
    cmdD5.Visible = False
    TotalC = TotalC + 1 'adds 1 to TotalC
End Sub

Private Sub cmdD1_Click()
    TotalA = 0 'gives TotalA an initial value of 0
    TotalB = 0 'sets initial value of TotalB to 0
    TotalC = 0 'sets initial value of TotalC to 0
    TotalD = 0 'sets initial value of TotalD to 0
    cmdD1.Visible = True
    cmdB1.Visible = False
    cmdC1.Visible = False
    cmdA1.Visible = False
    TotalD = TotalD + 1 'adds 1 to TotalD
End Sub

Private Sub cmdD2_Click()
    cmdA2.Visible = False
    cmdB2.Visible = False
    cmdC2.Visible = False
    cmdD2.Visible = True
    TotalD = TotalD + 1 'adds 1 to TotalD
End Sub

Private Sub cmdD3_Click()
    cmdA3.Visible = False
    cmdB3.Visible = False
    cmdC3.Visible = False
    cmdD3.Visible = True
    TotalD = TotalD + 1 'adds 1 to TotalD
End Sub

Private Sub cmdD4_Click()
    cmdA4.Visible = False
    cmdB4.Visible = False
    cmdC4.Visible = False
    cmdD4.Visible = True
    TotalD = TotalD + 1 'adds 1 to TotalD
End Sub

Private Sub cmdD5_Click()
    cmdA5.Visible = False
    cmdB5.Visible = False
    cmdC5.Visible = False
    cmdD5.Visible = True
    TotalD = TotalD + 1 'adds 1 to TotalD
End Sub

Private Sub cmdQuit_Click()
    End 'quits program
End Sub

Private Sub cmdReturn_Click()
    Form4.Hide 'hides Sorting page from user
    Form2.Show 'shows main page to user
End Sub

Private Sub cmdSort_Click()
    Dim Answer As String
    Dim Category As String
    picResults.Cls
        If (TotalA >= 2 And TotalB < 2 And TotalC < 2 And TotalD < 2) Or (TotalA >= 3 And TotalB <= 2 And TotalC <= 2 And TotalD <= 2) Then
        'if A has the majority of numbers (1-5), then the following message is displayed to the user:
            picResults.Print "You died an awful, awful death. Its"
            picResults.Print "because, you didn't set the right pace and"
            picResults.Print "you crossed the river at the wrong time of day."
            picResults.Print "If it weren't for your bad luck you woud have"
            picResults.Print "made it to Chimney Rock and met Buffalo Bill. Yet"
            picResults.Print "you became so ill your party fed."
            picResults.Print "you to a wild band of prarie dogs. "
            picResults2.Picture = LoadPicture(App.Path & "\PrairieDog.jpg")
        ElseIf (TotalB >= 2 And TotalA < 2 And TotalC < 2 And TotalD < 2) Or (TotalB >= 3 And TotalA <= 2 And TotalC <= 2 And TotalD <= 2) Then
       'if B has the majority of numbers (1-5), then the following message is displayed to the user:
            picResults.Print "You made it!  After crossing the river the proper way,"
            picResults.Print "rationing your food correctly and taking care of all your"
            picResults.Print "ailments in the right way, you arrived in Oregon with a"
            picResults.Print "perfect bill of health. However, after the journey you"
            picResults.Print "are an emotional wreck and refuse to speak."
            picResults.Print "Consider moving to California."
            picResults2.Picture = LoadPicture(App.Path & "\GoldMiner.jpg")
        ElseIf (TotalC >= 2 And TotalB < 2 And TotalA < 2 And TotalD < 2) Or (TotalC >= 3 And TotalB <= 2 And TotalA <= 2 And TotalD <= 2) Then
        'if C has the majority of numbers (1-5), then the following message is displayed to the user:
            picResults.Print "Congratulations! You almost made it! While trying"
            picResults.Print "to cross the river your party was swept downstream"
            picResults.Print "and you lost your wagon. Walking to Oregon,"
            picResults.Print "has been done before but not without considerable"
            picResults.Print "hardship. After walking for a few weeks you"
            picResults.Print "were gored by a disgruntled buffalo. Your"
            picResults.Print "party would have saved you but, they too were"
            picResults.Print "gored. Your hardships will be remembered."
            picResults2.Picture = LoadPicture(App.Path & "\Bison.jpg")
        ElseIf (TotalD >= 2 And TotalB < 2 And TotalC < 2 And TotalA < 2) Or (TotalD >= 3 And TotalB <= 2 And TotalC <= 2 And TotalA <= 2) Then
        'if D has the majority of numbers (1-5), then the following message is displayed to the user:
            picResults.Print "You ran into a nomadic tribe from Scotland."
            picResults.Print "Seeing that they had a noble cause"
            picResults.Print "you joined them and never returned to America"
            picResults2.Picture = LoadPicture(App.Path & "\braveheart.jpg")
        ElseIf (TotalA = TotalB) Then
            Category = InputBox("What is stronger animal: a bison or a buffalo", , "A tough choice.")
            'asks the user a tie-breaking question if TotalA is = to TotalB
            If LCase(Category) = "buffalo" Then
            'if the answer to the tie-breaking question is "buffalo" then the following message is displayed to the user:
                picResults.Print "So, you like totally manifested destiny. Here's to you"
                picResults.Print " Here's to you for pushing into the frontier."
                picResults.Print "If it weren't for you we wouldn't be the great"
                picResults.Print "country we are. Your actions echo the famous song"
                picResults.Print " by Woodie Guthrie'This Land is My Land, this land '."
                picResults.Print " is NOT your Land.' While those might not be the"
                picResults.Print "actual lyrics you get the drift. "
                picResults.Print "You civilized the West. Didnt' you?"
                picResults2.Picture = LoadPicture(App.Path & "\ManifestDestiny.jpg")
            Else
            'if the answer to the tie-breaking question is not "buffalo" then the following message is displayed to the user:
            picResults.Print "You died an awful, awful death. Its"
            picResults.Print "because, you didn't set the right pace and"
            picResults.Print "you crossed the river at the wrong time of day."
            picResults.Print "If it weren't for your bad luck you woud have"
            picResults.Print "made it to Chimney Rock"
            picResults.Print "Yet you became so ill your party fed."
            picResults.Print "you to a wild band of prarie dogs. "
            picResults2.Picture = LoadPicture(App.Path & "\PrairieDog.jpg")
            End If
        ElseIf (TotalA = TotalC) Then
            'asks the user a tie-breaking question if TotalA = TotalC
            Category = InputBox("Would you rather ride a bison or a unicorn", , "A hard choice!")
            If LCase(Category) = "unicorn" Then
            'if the answer to the tie-breaking question is "unicorn" then the following message is displayed to the user:
                picResults.Print "Unicorns don't exist! And that's why you died on"
                picResults.Print "on the Oregon Trail. If you had simply remembered what"
                picResults.Print "your grandmother taught you--that Unicorns can only."
                picResults.Print "be ridden by elves in Narnia-- you would have lived."
                picResults.Print "Nonetheless, the rest of your party made it to chimney rock,"
                picResults.Print "by the first day. They were however, crushed under a meteor"
                picResults.Print "that fell from the sky after one said 'God doesn't exist'."
                picResults2.Picture = LoadPicture(App.Path & "\Unicorn.jpg")
            Else
            'if the answer to the tie-breaking question is not "unicorn" then the following message is displayed to the user:
            picResults.Print "You died an awful, awful death. Its"
            picResults.Print "because, you didn't set the right pace and"
            picResults.Print "you crossed the river at the wrong time of day."
            picResults.Print "If it weren't for your bad luck you woud have"
            picResults.Print "made it to Chimney Rock and met Buffalo Bill. Yet"
            picResults.Print "you became so ill your party fed."
            picResults.Print "you to a wild band of prarie dogs. "
            picResults2.Picture = LoadPicture(App.Path & "\PrairieDog.jpg")
            End If
        ElseIf (TotalA = TotalD) Then
        'asks the user a tie-breaking question if TotalA = TotalD
            Category = InputBox("How cool is John Wayne?: Cool or So Uncool", , "Your Fate Rests in this Question")
            If LCase(Category) = "cool" Then
            'if the answer to the tie-breaking question is "cool" then the following message is displayed to the user:
            picResults.Print "You made it!  Cause John Way took away all the hardships.,"
            picResults.Print "He rationed your food correctly and took care of all your"
            picResults.Print "ailments in the right way. You arrived in Oregon with a"
            picResults.Print "perfect bill of health. However, you may now have a problem."
            picResults.Print "John Wayne wants your soul. He says it's payment for all."
            picResults.Print "He's done. Good luck getting out of grasp of Satan."
            picResults2.Picture = LoadPicture(App.Path & "\JohnWayne.jpg")
            Else
            'if the answer to the tie-breaking question is not "cool" then the following message is displayed to the user:
                picResults.Print "Let's just say you got shot by"
                picResults.Print "one sad lone ranger and the rest of"
                picResults.Print "your journey didn't go well either. Ma set fire to"
                picResults.Print "the wagon while trying to smoke her pipe and everyone"
                picResults.Print "perished."
                picResults.Print "You never made it Oregon."
                picResults2.Picture = LoadPicture(App.Path & "\JohnWayne.jpg")
            End If
        ElseIf (TotalB = TotalC) Then
        'asks the user a deciding question of TotalB = TotalC
            Category = InputBox("Who is cooler: Buffalo Bill or Sitting Bull?", , "Prove Your Knowledge")
        'if the answer to the tie-breaking question is "surprise" then the following message is displayed to the user:
            If LCase(Category) = "buffalo bill" Then
                picResults.Print "So you're really into slaughtering thousands of"
                picResults.Print "helpless animals in the name of 'Progress'? Well I"
                picResults.Print "guess that's why you got to Oregon. Being cut-throat"
                picResults.Print "is how you survive. Nonetheless, you will spend your"
                picResults.Print "next incarnation as a maggot."
                picResults2.Picture = LoadPicture(App.Path & "\BuffaloBill.jpg")
            Else
            'if the answer to the tie-breaking question is not "surprise" then the following message is displayed to the user:
                picResults.Print "Sitting Bull was assasinated by the United States"
                picResults.Print " after the battle of Little Big Horn."
                picResults.Print "Also known as Ta-tanka I-Yotank, he was a Lakota,"
                picResults.Print "Sioux Holy Man who participated in many struggles "
                picResults.Print "against Western Expansion, not unlike the Oregon"
                picResults.Print "trail. You didn't make it to Oregon because you"
                picResults.Print "realized Western Expansion was not for the benefit"
                picResults.Print "of everyone. Particulary folks like Sitting Bull"
                picResults2.Picture = LoadPicture(App.Path & "\SittingBull.jpg")
            End If
        ElseIf (TotalB = TotalD) Then
            'asks the user a tie-breaking question of TotalB = TotalD
            Category = InputBox("You must eat one of your party. Who should it be? Jeb or Jud", , "A tasy decision")
            If LCase(Category) = "jud" Then
            'if the answer to the tie-breaking question is "Jud" then the following message is displayed to the user:
                picResults.Print "Jud was a scrawny man but you were able to make him,"
                picResults.Print "last with your other rations. Crossing the river proved"
                picResults.Print "difficult for your party and you lost most of your supplies"
                picResults.Print "However, using the wit, grit, and spit of the pioneer"
                picResults.Print "spirit you were able to make it all the way to Chimney Rock"
                picResults.Print "whereupon you were abducted by aliens."
                picResults2.Picture = LoadPicture(App.Path & "\ChimneyRock.jpg")
            Else
            'if the answer to the tie-breaking question is not "Jud" then the following message is displayed to the user:
                picResults.Print "Jed was meaty fellow and he didn't go down,"
                picResults.Print "wihtout a fight. Too bad you chose to eat him"
                picResults.Print "him in the middle of a river crossing. Your whole"
                picResults.Print "party perished."
                picResults2.Picture = LoadPicture(App.Path & "\RiverCrossing.jpg")
            End If
        ElseIf (TotalC = TotalD) Then
        'asks the user a tie-breaking question of TotalD = TotalC
        Category = InputBox("Are you traveling West: Yes or No", , "Are You Traveling West")
            If LCase(Category) = "Yes" Then
            'if the answer to the tie-breaking question is "Yes" then the following message is displayed to the user:
                picResults.Print "It's a good thing you're traveling West, because"
                picResults.Print "anywhere else and you wouldn't got to sacred"
                picResults.Print "Oregon. Your decisions proved fruitful and you were"
                picResults.Print "able to make it despite a few mishap. Ma perished"
                picResults.Print "in route and Pa realized he forgot his favorite comb"
                picResults.Print "and walked back to Minnesota. But well done! You"
                picResults.Print "made it to Oregon. All the other parties died."
                picResults2.Picture = LoadPicture(App.Path & "\WestWardExpansion.jpg")
            Else
            'if the answer to the tie-breaking question is not "Yes" then the following message is displayed to the user:
                picResults.Print "You're not traveling West? That may explain why you"
                picResults.Print "arrived at a nice resort in Mexico. No need to worry"
                picResults.Print "about all the misfortunes along the Oregon trail."
                picResults.Print "Sit back and enjoy the pool side view."
                picResults2.Picture = LoadPicture(App.Path & "\PoolSide.jpg")
            End If
    End If 'the following Visible commands reset the buttons on the sorting page so user can play again
    cmdA1.Visible = True
    cmdB1.Visible = True
    cmdC1.Visible = True
    cmdD1.Visible = True
    cmdA2.Visible = True
    cmdB2.Visible = True
    cmdC2.Visible = True
    cmdD2.Visible = True
    cmdA3.Visible = True
    cmdB3.Visible = True
    cmdC3.Visible = True
    cmdD3.Visible = True
    cmdA4.Visible = True
    cmdB4.Visible = True
    cmdC4.Visible = True
    cmdD4.Visible = True
    cmdA5.Visible = True
    cmdB5.Visible = True
    cmdC5.Visible = True
    cmdD5.Visible = True
    'the following four commands reset the total values for A - D to 0 so user can play again
    TotalA = 0
    TotalB = 0
    TotalC = 0
    TotalD = 0
End Sub

Private Sub Command1_Click()

    Form4.Hide
    frmhuntingone.Show


End Sub

Private Sub Form_Activate()
    picResults2.Picture = LoadPicture("") 'resets picResults 2 to an empty display to user can play again
End Sub



Private Sub picHat_Click()

End Sub

