VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   9360
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
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
         Name            =   "Mistral"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   2775
      Left            =   6120
      ScaleHeight     =   2715
      ScaleWidth      =   4635
      TabIndex        =   20
      Top             =   600
      Width           =   4695
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
      Left            =   600
      TabIndex        =   19
      Top             =   5520
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
      Left            =   6840
      ScaleHeight     =   4275
      ScaleWidth      =   3195
      TabIndex        =   18
      Top             =   3480
      Width           =   3255
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
      Caption         =   "How Will You Decide to Travel"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   6975
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   3375
      Left            =   360
      Top             =   5280
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Left            =   6000
      Top             =   600
      Width           =   4935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   4335
      Left            =   6720
      Top             =   3480
      Width           =   3495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this is the game part of Oregon Trail

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
    frm4.Hide 'hides Sorting page from user
    frm2.Show 'shows main page to user
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
            picResults.Print "If it weren't for your tenacity you woudn't have"
            picResults.Print "made it to Chimney Rock and met Buffalo Bill. Yet"
            picResults.Print "after Chimney Rock you became so ill your party fed."
            picResults.Print "you to a wild band of prarie dogs. Oh well. Se La Vi."
            picResults2.Picture = LoadPicture(App.Path & "\PrairieDog.jpg")
        ElseIf (TotalB >= 2 And TotalA < 2 And TotalC < 2 And TotalD < 2) Or (TotalB >= 3 And TotalA <= 2 And TotalC <= 2 And TotalD <= 2) Then
       'if B has the majority of numbers (1-5), then the following message is displayed to the user:
            picResults.Print "You made it!  After crossing the river the proper way,"
            picResults.Print "rationing your food correctly and taking care of all your"
            picResults.Print "ailments in the right way, you arrived in Oregon with a"
            picResults.Print "perfect bill of health. However, after the journey you"
            picResults.Print "are an emotional wreck and refuse to speak."
            picResults.Print "Consider moving to California."
            picResults2.Picture = LoadPicture(App.Path & "\SortRavenclaw.jpg")
        ElseIf (TotalC >= 2 And TotalB < 2 And TotalA < 2 And TotalD < 2) Or (TotalC >= 3 And TotalB <= 2 And TotalA <= 2 And TotalD <= 2) Then
        'if C has the majority of numbers (1-5), then the following message is displayed to the user:
            picResults.Print "You are a Hufflepuff!  Hufflepuff, founded by"
            picResults.Print "Helga Hufflepuff, is the most inclusive among"
            picResults.Print "the four houses, valuing hard work, loyalty,"
            picResults.Print "determination, patience, friendship and fair"
            picResults.Print "play rather than a particular aptitude in its"
            picResults.Print "members.  Its animal is the badger, and yellow"
            picResults.Print "and black are its colours. Pomona Sprout is the"
            picResults.Print "the Head of House. The Fat Friar is its ghost."
            picResults2.Picture = LoadPicture(App.Path & "\SortHufflepuff.jpg")
        ElseIf (TotalD >= 2 And TotalB < 2 And TotalC < 2 And TotalA < 2) Or (TotalD >= 3 And TotalB <= 2 And TotalC <= 2 And TotalA <= 2) Then
        'if D has the majority of numbers (1-5), then the following message is displayed to the user:
            picResults.Print "You are a Gryffindor!  Gryffindor values courage,"
            picResults.Print "chivalry and boldness. Its animal is the lion, and"
            picResults.Print "its colors are scarlet and gold.  Minerva Mcgonagall"
            picResults.Print "is the most recent Head of House.  Nearly Headless"
            picResults.Print "Nick is the house ghost. The founder of Gryffindor"
            picResults.Print "is Godric Gryffindor."
            picResults2.Picture = LoadPicture(App.Path & "\SortGryffindor.jpg")
        ElseIf (TotalA = TotalB) Then
            Category = InputBox("Would you rather be given a locket or a surprise?", , "Please pick a color.")
            'asks the user a tie-breaking question if TotalA is = to TotalB
            If LCase(Category) = "locket" Then
            'if the answer to the tie-breaking question is "locket" then the following message is displayed to the user:
                picResults.Print "You are a Slytherin!  Like Salazar Slytherin, its"
                picResults.Print "founder, Slytherin House values ambition, cunning"
                picResults.Print "and resourcefulness.  Its emblematic animal is the"
                picResults.Print "serpent, and its colours are green and silver."
                picResults.Print "Professor Horace Slughorn is the Head of Slytherin,"
                picResults.Print "replacing Severus Snape who has recently left"
                picResults.Print "Hogwarts.  The Bloody Baron is the house ghost."
                picResults2.Picture = LoadPicture(App.Path & "\SortSlytherin.jpg")
            Else
            'if the answer to the tie-breaking question is not "locket" then the following message is displayed to the user:
                picResults.Print "You are a Ravenclaw!  Ravenclaw values intelligence,"
                picResults.Print "knowledge, and wit. Its animal is the eagle, and its"
                picResults.Print "colors are blue and bronze.  The Ravenclaw Head of"
                picResults.Print "House is Filius Flitwich, and the Ravenclaw ghost"
                picResults.Print "is the Grey Lady. The house was founded by Rowena"
                picResults.Print "Ravenclaw."
                picResults2.Picture = LoadPicture(App.Path & "\SortRavenclaw.jpg")
            End If
        ElseIf (TotalA = TotalC) Then
            'asks the user a tie-breaking question if TotalA = TotalC
            Category = InputBox("Would you rather be given a locket or a goblet?", , "Tie-breaker!")
            If LCase(Category) = "locket" Then
            'if the answer to the tie-breaking question is "locket" then the following message is displayed to the user:
                picResults.Print "You are a Slytherin!  Like Salazar Slytherin, its"
                picResults.Print "founder, Slytherin House values ambition, cunning"
                picResults.Print "and resourcefulness.  Its emblematic animal is the"
                picResults.Print "serpent, and its colours are green and silver."
                picResults.Print "Professor Horace Slughorn is the Head of Slytherin,"
                picResults.Print "replacing Severus Snape who has recently left"
                picResults.Print "Hogwarts.  The Bloody Baron is the house ghost."
                picResults2.Picture = LoadPicture(App.Path & "\SortSlytherin.jpg")
            Else
            'if the answer to the tie-breaking question is not "locket" then the following message is displayed to the user:
                picResults.Print "You are a Hufflepuff!  Hufflepuff, founded by"
                picResults.Print "Helga Hufflepuff, is the most inclusive among"
                picResults.Print "the four houses, valuing hard work, loyalty,"
                picResults.Print "determination, patience, friendship and fair"
                picResults.Print "play rather than a particular aptitude in its"
                picResults.Print "members.  Its animal is the badger, and yellow"
                picResults.Print "and black are its colours. Pomona Sprout is the"
                picResults.Print "the Head of House. The Fat Friar is its ghost."
                picResults2.Picture = LoadPicture(App.Path & "\SortHufflepuff.jpg")
            End If
        ElseIf (TotalA = TotalD) Then
        'asks the user a tie-breaking question if TotalA = TotalD
            Category = InputBox("Would you rather be given a locket or a sword?", , "Tie-breaker!")
            If LCase(Category) = "locket" Then
            'if the answer to the tie-breaking question is "locket" then the following message is displayed to the user:
                picResults.Print "You are a Slytherin!  Like Salazar Slytherin, its"
                picResults.Print "founder, Slytherin House values ambition, cunning"
                picResults.Print "and resourcefulness.  Its emblematic animal is the"
                picResults.Print "serpent, and its colours are green and silver."
                picResults.Print "Professor Horace Slughorn is the Head of Slytherin,"
                picResults.Print "replacing Severus Snape who has recently left"
                picResults.Print "Hogwarts.  The Bloody Baron is the house ghost."
                picResults2.Picture = LoadPicture(App.Path & "\SortSlytherin.jpg")
            Else
            'if the answer to the tie-breaking question is not "locket" then the following message is displayed to the user:
                picResults.Print "You are a Gryffindor!  Gryffindor values courage,"
                picResults.Print "chivalry and boldness. Its animal is the lion, and"
                picResults.Print "its colors are scarlet and gold.  Minerva Mcgonagall"
                picResults.Print "is the most recent Head of House.  Nearly Headless"
                picResults.Print "Nick is the house ghost. The founder of Gryffindor"
                picResults.Print "is Godric Gryffindor."
                picResults2.Picture = LoadPicture(App.Path & "\SortGryffindor.jpg")
            End If
        ElseIf (TotalB = TotalC) Then
        'asks the user a tie-breaking question of TotalB = TotalC
            Category = InputBox("Would you rather be given a surprise or a goblet?", , "Tie-breaker!")
        'if the answer to the tie-breaking question is "surprise" then the following message is displayed to the user:
            If LCase(Category) = "surprise" Then
                picResults.Print "You are a Ravenclaw!  Ravenclaw values intelligence,"
                picResults.Print "knowledge, and wit. Its animal is the eagle, and its"
                picResults.Print "colors are blue and bronze.  The Ravenclaw Head of"
                picResults.Print "House is Filius Flitwich, and the Ravenclaw ghost"
                picResults.Print "is the Grey Lady. The house was founded by Rowena"
                picResults.Print "Ravenclaw."
                picResults2.Picture = LoadPicture(App.Path & "\SortRavenclaw.jpg")
            Else
            'if the answer to the tie-breaking question is not "surprise" then the following message is displayed to the user:
                picResults.Print "You are a Hufflepuff!  Hufflepuff, founded by"
                picResults.Print "Helga Hufflepuff, is the most inclusive among"
                picResults.Print "the four houses, valuing hard work, loyalty,"
                picResults.Print "determination, patience, friendship and fair"
                picResults.Print "play rather than a particular aptitude in its"
                picResults.Print "members.  Its animal is the badger, and yellow"
                picResults.Print "and black are its colours. Pomona Sprout is the"
                picResults.Print "the Head of House. The Fat Friar is its ghost."
                picResults2.Picture = LoadPicture(App.Path & "\SortHufflepuff.jpg")
            End If
        ElseIf (TotalB = TotalD) Then
            'asks the user a tie-breaking question of TotalB = TotalD
            Category = InputBox("Would you rather be given a surprise or a sword?", , "Tie-breaker!")
            If LCase(Category) = "surprise" Then
            'if the answer to the tie-breaking question is "surprise" then the following message is displayed to the user:
                picResults.Print "You are a Ravenclaw!  Ravenclaw values intelligence,"
                picResults.Print "knowledge, and wit. Its animal is the eagle, and its"
                picResults.Print "colors are blue and bronze.  The Ravenclaw Head of"
                picResults.Print "House is Filius Flitwich, and the Ravenclaw ghost"
                picResults.Print "is the Grey Lady. The house was founded by Rowena"
                picResults.Print "Ravenclaw."
                picResults2.Picture = LoadPicture(App.Path & "\SortRavenclaw.jpg")
            Else
            'if the answer to the tie-breaking question is not "surprise" then the following message is displayed to the user:
                picResults.Print "You are a Gryffindor!  Gryffindor values courage,"
                picResults.Print "chivalry and boldness. Its animal is the lion, and"
                picResults.Print "its colors are scarlet and gold.  Minerva Mcgonagall"
                picResults.Print "is the most recent Head of House.  Nearly Headless"
                picResults.Print "Nick is the house ghost. The founder of Gryffindor"
                picResults.Print "is Godric Gryffindor."
                picResults2.Picture = LoadPicture(App.Path & "\SortGryffindor.jpg")
            End If
        ElseIf (TotalC = TotalD) Then
        'asks the user a tie-breaking question of TotalD = TotalC
        Category = InputBox("Would you rather be given a goblet or a sword?", , "Tie-breaker!")
            If LCase(Category) = "goblet" Then
            'if the answer to the tie-breaking question is "goblet" then the following message is displayed to the user:
                picResults.Print "You are a Hufflepuff!  Hufflepuff, founded by"
                picResults.Print "Helga Hufflepuff, is the most inclusive among"
                picResults.Print "the four houses, valuing hard work, loyalty,"
                picResults.Print "determination, patience, friendship and fair"
                picResults.Print "play rather than a particular aptitude in its"
                picResults.Print "members.  Its animal is the badger, and yellow"
                picResults.Print "and black are its colours. Pomona Sprout is the"
                picResults.Print "the Head of House. The Fat Friar is its ghost."
                picResults2.Picture = LoadPicture(App.Path & "\SortHufflepuff.jpg")
            Else
            'if the answer to the tie-breaking question is not "goblet" then the following message is displayed to the user:
                picResults.Print "You are a Gryffindor!  Gryffindor values courage,"
                picResults.Print "chivalry and boldness. Its animal is the lion, and"
                picResults.Print "its colors are scarlet and gold.  Minerva Mcgonagall"
                picResults.Print "is the most recent Head of House.  Nearly Headless"
                picResults.Print "Nick is the house ghost. The founder of Gryffindor"
                picResults.Print "is Godric Gryffindor."
                picResults2.Picture = LoadPicture(App.Path & "\SortGryffindor.jpg")
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

Private Sub Form_Activate()
    picResults2.Picture = LoadPicture("") 'resets picResults 2 to an empty display to user can play again
End Sub



Private Sub picHat_Click()

End Sub
