VERSION 5.00
Begin VB.Form frmEast 
   BackColor       =   &H000080FF&
   Caption         =   "East Regional"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF0000&
      Height          =   1695
      Left            =   6720
      ScaleHeight     =   1635
      ScaleWidth      =   3075
      TabIndex        =   39
      Top             =   2640
      Width           =   3135
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PLEASE CLICK!!! ON WINNER OF REGION TO SUBMIT THAT WINNER TO FINAL FOUR BRACKET"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   7680
      Picture         =   "frmEast.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   37
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdFinalFour 
      Caption         =   "Go To Final Four"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   35
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton CmdGoToSouthBracket 
      Caption         =   "Go To South Bracket"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   34
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoToWest 
      Caption         =   "Go To West Bracket"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   33
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Section Total"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   32
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdGoToMidwest 
      Caption         =   "Go To Midwest Bracket"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   31
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdEastWinner 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner13 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner14 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner12 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner11 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner10 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner8 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner7 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner6 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner5 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner4 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner3 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdBelmont 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15 Belmont"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdGeorgetown 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2 Georgetown"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdTexasTech 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10 Texas Tech"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdBostonCollege 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7 Boston College"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdOralRoberts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14 Oral Roberts"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdWashingtonSt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3 Washington St"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdGWashington 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11 G Washington"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdVanderbilt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6 Vanderbilt"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewMexicoSt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13 New Mexico St"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdTexas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4 Texas"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdArkansas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12 Arkansas"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdUSC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5 USC"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner9 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdMichiganSt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9 Michigan St"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdMarquette 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8 Marquette"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdEasternKy 
      BackColor       =   &H00FFFFFF&
      Caption         =   "16 Eastern Ky"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdNorthCarolina 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1 North Carolina"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   2295
      Left            =   7080
      ScaleHeight     =   2235
      ScaleWidth      =   3675
      TabIndex        =   38
      Top             =   5880
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "TO SEE YOUR SCORES CLICK!! ON SECTION TOTAL TO TRANSFER SCORE TO TOTALS PAGE AND SUBMIT BRACKET"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4920
      TabIndex        =   41
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   8400
      Picture         =   "frmEast.frx":914D
      Top             =   4080
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "East Regional Bracket"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   36
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmEast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'This form is where the user will begin to start and select who they believe will win each round
' this form specifically is for the east bracket

Private Sub cmdCompute_Click()
    Dim EastR1(1 To 8) As String                            'array for east round 1 team names
    Dim EastR1Pos(1 To 8) As Integer                        'this array is for use if we decide to do scoring where rankings of teams are neede and we want to do upsets are worth more
    Dim EastR2(1 To 4) As String                            'the other arrays are the same except for rounds 2 and 3
    Dim EastR2Pos(1 To 4) As Integer
    Dim EastR3(1 To 2) As String
    Dim EastR3Pos(1 To 2) As Integer
    Dim Ctr As Integer, CTR2 As Integer, CTR3 As Integer    'counter for each array
    Open App.Path & "\EastR1.txt" For Input As #1           'open notepad with round 1 winners
    Ctr = 0
    Do Until EOF(1)                                         'loop to set ctr to EastR1.txt
        Ctr = Ctr + 1
        Input #1, EastR1Pos(Ctr), EastR1(Ctr)
    Loop
    Close #1
    
    EastR1Sum = 0                                           'set sum to 0
    If cmdwinner1.Caption = EastR1(1) Then                  'The following are all if statements to add a point to the users score if the if statement is true
        EastR1Sum = EastR1Sum + 1
    End If
    If cmdwinner2.Caption = EastR1(2) Then
        EastR1Sum = EastR1Sum + 1
    End If
    If cmdwinner3.Caption = EastR1(3) Then
        EastR1Sum = EastR1Sum + 1
    End If
    If cmdwinner4.Caption = EastR1(4) Then
        EastR1Sum = EastR1Sum + 1
    End If
    If cmdwinner5.Caption = EastR1(5) Then
        EastR1Sum = EastR1Sum + 1
    End If
    If cmdwinner6.Caption = EastR1(6) Then
        EastR1Sum = EastR1Sum + 1
    End If
    If cmdwinner7.Caption = EastR1(7) Then
        EastR1Sum = EastR1Sum + 1
    End If
    If cmdwinner8.Caption = EastR1(8) Then
        EastR1Sum = EastR1Sum + 1
    End If
    
    Open App.Path & "\EastR2.txt" For Input As #2                 'opening the second notepade for winners of east round 2
    CTR2 = 0
    Do Until EOF(2)                                                'setting counter to round 2 teams
        CTR2 = CTR2 + 1
        Input #2, EastR2Pos(CTR2), EastR2(CTR2)
    Loop
    Close #2
    
    EastR2Sum = 0
    If cmdwinner9.Caption = EastR2(1) Then                          'all if statments are adding 2 to the users total if the statement is true
        EastR2Sum = EastR2Sum + 2
    End If
    If cmdwinner10.Caption = EastR2(2) Then
        EastR2Sum = EastR2Sum + 2
    End If
    If cmdwinner11.Caption = EastR2(3) Then
        EastR2Sum = EastR2Sum + 2
    End If
    If cmdwinner12.Caption = EastR2(4) Then
        EastR2Sum = EastR2Sum + 2
    End If
    
    Open App.Path & "\EastR3.txt" For Input As #3                   'opening the third notepade for winners of east round 3
    CTR3 = 0
    Do Until EOF(3)                                                    'setting counter to round 3 teams
        CTR3 = CTR3 + 1
        Input #3, EastR3Pos(CTR3), EastR3(CTR3)
    Loop
    Close #3
    
    EastR3Sum = 0
    If cmdwinner13.Caption = EastR3(1) Then                         'all if statments are adding 4 to the users total if the statement is true
        EastR3Sum = EastR3Sum + 4
    End If
    If cmdwinner14.Caption = EastR3(2) Then
        EastR3Sum = EastR3Sum + 4
    End If
    
    EastTotal = EastR1Sum + EastR2Sum + EastR3Sum
End Sub

'taking the caption from this button and setting it equal to EastWinner to later have the caption show up on final four form
Private Sub cmdEastWinner_Click()
    EastWinner = cmdEastWinner.Caption
End Sub


'The following buttons will allow the user to switch from form to form while they choose there teams and eventually will move to the final four form
Private Sub cmdFinalFour_Click()
    frmEast.Hide
    frmFinals.Show
End Sub

Private Sub cmdGoToMidwest_Click()
    frmEast.Hide
    frmMidwest.Show
End Sub

Private Sub CmdGoToSouthBracket_Click()
    frmEast.Hide
    frmSouth.Show
End Sub

Private Sub cmdGoToWest_Click()
    frmEast.Hide
    frmWest.Show
End Sub

'all of the following allow the user to click on a button and transfer the caption of the button to the next round button
'This allows the user to pick there teams and see who they are picking
Private Sub cmdNorthCarolina_Click()
    cmdwinner1.Caption = "1 North Carolina"
End Sub
Private Sub cmdEasternKy_Click()
    cmdwinner1.Caption = "16 Eastern Ky"
End Sub
Private Sub cmdMarquette_Click()
    cmdwinner2.Caption = "8 Marquette"
End Sub
Private Sub cmdMichiganSt_Click()
    cmdwinner2.Caption = "9 Michigan St"
End Sub
Private Sub cmdUSC_Click()
    cmdwinner3.Caption = "5 USC"
End Sub
Private Sub cmdArkansas_Click()
    cmdwinner3.Caption = "12 Arkansas"
End Sub
Private Sub cmdTexas_Click()
    cmdwinner4.Caption = "4 Texas"
End Sub
Private Sub cmdNewMexicoSt_Click()
    cmdwinner4.Caption = "13 New Mexico St"
End Sub
Private Sub cmdVanderbilt_Click()
    cmdwinner5.Caption = "6 Vanderbilt"
End Sub
Private Sub cmdGWashington_Click()
    cmdwinner5.Caption = "11 GWashington"
End Sub
Private Sub cmdWashingtonSt_Click()
    cmdwinner6.Caption = "3 Washington St"
End Sub
Private Sub cmdOralRoberts_Click()
    cmdwinner6.Caption = "14 Oral Roberts"
End Sub
Private Sub cmdBostonCollege_Click()
    cmdwinner7.Caption = "7 Boston College"
End Sub
Private Sub cmdTexasTech_Click()
    cmdwinner7.Caption = "10 Texas Tech"
End Sub
Private Sub cmdGeorgetown_Click()
    cmdwinner8.Caption = "2 Georgetown"
End Sub
Private Sub cmdBelmont_Click()
    cmdwinner8.Caption = "15 Belmont"
End Sub

Private Sub cmdwinner1_Click()
    cmdwinner9.Caption = cmdwinner1.Caption
End Sub
Private Sub cmdwinner2_Click()
    cmdwinner9.Caption = cmdwinner2.Caption
End Sub
Private Sub cmdwinner3_Click()
    cmdwinner10.Caption = cmdwinner3.Caption
End Sub
Private Sub cmdwinner4_Click()
    cmdwinner10.Caption = cmdwinner4.Caption
End Sub
Private Sub cmdwinner5_Click()
    cmdwinner11.Caption = cmdwinner5.Caption
End Sub
Private Sub cmdwinner6_Click()
    cmdwinner11.Caption = cmdwinner6.Caption
End Sub
Private Sub cmdwinner7_Click()
    cmdwinner12.Caption = cmdwinner7.Caption
End Sub
Private Sub cmdwinner8_Click()
    cmdwinner12.Caption = cmdwinner8.Caption
End Sub
Private Sub cmdwinner9_Click()
   cmdwinner13.Caption = cmdwinner9.Caption
End Sub
Private Sub cmdwinner10_Click()
    cmdwinner13.Caption = cmdwinner10.Caption
End Sub
Private Sub cmdwinner11_Click()
    cmdwinner14.Caption = cmdwinner11.Caption
End Sub
Private Sub cmdwinner12_Click()
    cmdwinner14.Caption = cmdwinner12.Caption
End Sub
Private Sub cmdwinner13_Click()
    cmdEastWinner.Caption = cmdwinner13.Caption
End Sub
Private Sub cmdwinner14_Click()
    cmdEastWinner.Caption = cmdwinner14.Caption
End Sub

