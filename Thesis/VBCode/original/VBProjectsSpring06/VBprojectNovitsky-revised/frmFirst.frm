VERSION 5.00
Begin VB.Form frmFirst 
   BackColor       =   &H000000FF&
   Caption         =   "We need the Info!"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrFour 
      Interval        =   20000
      Left            =   6840
      Top             =   5160
   End
   Begin VB.Timer tmrThree 
      Interval        =   13000
      Left            =   6960
      Top             =   4200
   End
   Begin VB.Timer tmrTwo 
      Interval        =   1000
      Left            =   7560
      Top             =   4800
   End
   Begin VB.Timer tmrOne 
      Interval        =   7000
      Left            =   6360
      Top             =   4560
   End
   Begin VB.PictureBox PicDisplay 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      ScaleHeight     =   1335
      ScaleWidth      =   4215
      TabIndex        =   18
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox txtforty 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Text            =   "0"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtBench 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Text            =   "0"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtMile 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Text            =   "0"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtLeap 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Text            =   "0"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox txtShuttle 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Text            =   "0"
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtCoor 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Text            =   "0"
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdCoor 
      BackColor       =   &H0000FFFF&
      Caption         =   "Get Score!"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   4560
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Run Calculator!"
      Height          =   1215
      Left            =   2280
      TabIndex        =   2
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblSpeed 
      BackColor       =   &H000000FF&
      Caption         =   "Please insert your 40m time (sec)"
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblBench 
      BackColor       =   &H000000FF&
      Caption         =   "Please insert your bench press weight (lbs)"
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblMile 
      BackColor       =   &H000000FF&
      Caption         =   "Insert mile time (minutes with decimal for sec)"
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblJump 
      BackColor       =   &H000000FF&
      Caption         =   "Please insert vertical leap (meters)"
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblShuttle 
      BackColor       =   &H000000FF&
      Caption         =   "Please Insert your shuttle time (sec)"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblCoordination 
      BackColor       =   &H000000FF&
      Caption         =   "Please insert your hand to eye coordination score (to calculate click get score button!)"
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image imgSport 
      Height          =   2460
      Left            =   5040
      Picture         =   "frmFirst.frx":0000
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lblName 
      BackColor       =   &H000000FF&
      Caption         =   "Whats your name?"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click() 'opens about page
    frmFirst.Hide
    frmAbout.Show
End Sub

Private Sub cmdCalculate_Click() 'opens frm second/runs all calulations and inputs into arrays.
    Dim pos, pass, size As Integer
    Dim yes As String
    Dim forty, shuttle, mile, bench, vleap, handeye, temp As Double
    Dim tempname As String
    size = 14
    temp = 0
    forty = 0
    shuttle = 0
    vleap = 0
    bench = 0
    mile = 0
    handeye = 0
    person = txtName.Text 'loads user input into variable
    forty = Val(txtforty.Text) 'loads user input into variable
    shuttle = Val(txtShuttle.Text) 'loads user input into variable
    bench = Val(txtBench.Text) 'loads user input into variable
    mile = Val(txtMile.Text) 'loads user input into variable
    vleap = Val(txtLeap.Text) 'loads user input into variable
    handeye = Val(txtCoor.Text) 'loads user input into variable
    yes = InputBox("Are you sure you're done " & person & "? Beacause theres no going back! yes=coontinue, anything else means NO! (caps sensitive)", "Continue?", "yes") ' checks to insure the want to move onto next page
    If yes = "yes" Then 'beings loading the nessicary file and running calculations
        Open App.Path & "\sports.txt" For Input As #1
        pos = 0
        Do Until EOF(1)
            pos = pos + 1
            Input #1, arraysport(pos), arrayforty(pos), arrayshuttle(pos), arraybench(pos), arraymile(pos), arrayleap(pos), arrayhandeye(pos) 'adds data from file to repective arrays
        Loop
        pos = 0
        Close #1
        For pos = 1 To 14   'assigns values to user input through comparrison
            If forty <= arrayforty(pos) Then
                fortyscore(pos) = 10
            ElseIf (forty) <= arrayforty(pos) * 1.05 Then
                fortyscore(pos) = 9
            ElseIf (forty) <= arrayforty(pos) * 1.1 Then
                fortyscore(pos) = 8
            ElseIf (forty) <= arrayforty(pos) * 1.15 Then
                fortyscore(pos) = 7
            ElseIf (forty) <= arrayforty(pos) * 1.2 Then
                fortyscore(pos) = 6
            ElseIf forty <= arrayforty(pos) * 1.25 Then
                fortyscore(pos) = 5
            ElseIf forty <= arrayforty(pos) * 1.3 Then
                fortyscore(pos) = 4
            ElseIf forty <= arrayforty(pos) * 1.35 Then
                fortyscore(pos) = 3
            ElseIf forty <= arrayforty(pos) * 1.4 Then
                fortyscore(pos) = 2
            ElseIf forty <= arrayforty(pos) * 1.45 Then
                fortyscore(pos) = 1
            Else
                fortyscore(pos) = 0
            End If
        Next pos
        For pos = 1 To 14 'assigns values to user input through comparrison
            If shuttle <= arrayshuttle(pos) Then
                shuttlescore(pos) = 10
            ElseIf shuttle <= arrayshuttle(pos) * 1.05 Then
                shuttlescore(pos) = 9
            ElseIf shuttle <= arrayshuttle(pos) * 1.1 Then
                shuttlescore(pos) = 8
            ElseIf shuttle <= arrayshuttle(pos) * 1.15 Then
                shuttlescore(pos) = 7
            ElseIf shuttle <= arrayshuttle(pos) * 1.2 Then
                shuttlescore(pos) = 6
            ElseIf shuttle <= arrayshuttle(pos) * 1.25 Then
                shuttlescore(pos) = 5
            ElseIf shuttle <= arrayshuttle(pos) * 1.3 Then
                shuttlescore(pos) = 4
            ElseIf shuttle <= arrayshuttle(pos) * 1.35 Then
                shuttlescore(pos) = 3
            ElseIf shuttle <= arrayshuttle(pos) * 1.4 Then
                shuttlescore(pos) = 2
            ElseIf shuttle <= arrayshuttle(pos) * 1.45 Then
                shuttlescore(pos) = 1
            Else
                shuttlescore(pos) = 0
        End If
    Next pos
    For pos = 1 To 14 'assigns values to user input through comparrison
            If bench >= arraybench(pos) Then
                benchscore(pos) = 10
            ElseIf bench >= arraybench(pos) / 1.05 Then
                benchscore(pos) = 9
            ElseIf bench >= arraybench(pos) / 1.1 Then
                shuttlescore(pos) = 8
            ElseIf bench >= arraybench(pos) / 1.15 Then
                benchscore(pos) = 7
            ElseIf bench >= arraybench(pos) / 1.2 Then
                benchscore(pos) = 6
            ElseIf bench >= arraybench(pos) / 1.25 Then
                benchscore(pos) = 5
            ElseIf bench >= arraybench(pos) / 1.3 Then
                benchscore(pos) = 4
            ElseIf bench >= arraybench(pos) / 1.35 Then
                benchscore(pos) = 3
            ElseIf bench >= arraybench(pos) / 1.4 Then
                benchscore(pos) = 2
            ElseIf bench >= arraybench(pos) / 1.45 Then
                benchscore(pos) = 1
             Else
                benchscore(pos) = 0
            End If
        Next pos
        For pos = 1 To 14 'assigns values to user input through comparrison
            If mile <= arraymile(pos) Then
                milescore(pos) = 10
            ElseIf mile <= arraymile(pos) * 1.05 Then
                milescore(pos) = 9
            ElseIf mile <= arraymile(pos) * 1.1 Then
                milescore(pos) = 8
            ElseIf mile <= arraymile(pos) * 1.15 Then
                milescore(pos) = 7
            ElseIf mile <= arraymile(pos) * 1.2 Then
                milescore(pos) = 6
            ElseIf mile <= arraymile(pos) * 1.25 Then
                milescore(pos) = 5
            ElseIf mile <= arraymile(pos) * 1.3 Then
                milescore(pos) = 4
            ElseIf mile <= arraymile(pos) * 1.35 Then
                milescore(pos) = 3
            ElseIf mile <= arraymile(pos) * 1.4 Then
                milescore(pos) = 2
            ElseIf mile <= arraymile(pos) * 1.45 Then
                milescore(pos) = 1
             Else
                milescore(pos) = 0
            End If
        Next pos
            For pos = 1 To 14 'assigns values to user input through comparrison
            If vleap >= arrayleap(pos) Then
                vleapscore(pos) = 10
            ElseIf vleap >= arrayleap(pos) / 1.05 Then
                vleapscore(pos) = 9
            ElseIf vleap >= arrayleap(pos) / 1.1 Then
                vleapscore(pos) = 8
            ElseIf vleap >= arrayleap(pos) / 1.15 Then
               vleapscore(pos) = 7
            ElseIf vleap >= arrayleap(pos) / 1.2 Then
                vleapscore(pos) = 6
            ElseIf vleap >= arrayleap(pos) / 1.25 Then
                vleapscore(pos) = 5
            ElseIf vleap >= arrayleap(pos) / 1.3 Then
                vleapscore(pos) = 4
            ElseIf vleap >= arrayleap(pos) / 1.35 Then
                vleapscore(pos) = 3
            ElseIf vleap >= arrayleap(pos) / 1.4 Then
                vleapscore(pos) = 2
            ElseIf vleap >= arrayleap(pos) / 1.45 Then
                vleapscore(pos) = 1
             Else
                vleapscore(pos) = 0
            End If
        Next pos
        For pos = 1 To 14 'assigns values to user input through comparrison
            If handeye >= arrayhandeye(pos) Then
                handeyescore(pos) = 10
            ElseIf handeye >= arrayhandeye(pos) / 1.1 Then
                handeyescore(pos) = 9
            ElseIf handeye >= arrayhandeye(pos) / 1.2 Then
                handeyescore(pos) = 8
            ElseIf handeye >= arrayhandeye(pos) / 1.3 Then
               handeyescore(pos) = 7
            ElseIf handeye >= arrayhandeye(pos) / 1.4 Then
                handeyescore(pos) = 6
            ElseIf handeye >= arrayhandeye(pos) / 1.5 Then
                handeyescore(pos) = 5
            ElseIf handeye >= arrayhandeye(pos) / 1.6 Then
                handeyescore(pos) = 4
            ElseIf handeye >= arrayhandeye(pos) / 1.7 Then
                handeyescore(pos) = 3
            ElseIf handeye >= arrayhandeye(pos) / 1.8 Then
                handeyescore(pos) = 2
            ElseIf handeye >= arrayhandeye(pos) / 1.9 Then
                handeyescore(pos) = 1
             Else
                handeyescore(pos) = 0
            End If
        Next pos
            ' totals all of the input scores and preps them for sorting and comparison
            total(1) = benchscore(1) / 1.25 + shuttlescore(1) * 1.4 + fortyscore(1) * 1.6 + milescore(1) * 1.1 + vleapscore(1) + handeyescore(1) * 1.2
            total(2) = benchscore(2) / 1.75 + shuttlescore(2) * 1.5 + fortyscore(2) * 1.25 + milescore(2) + vleapscore(2) * 2 + handeyescore(2) * 1.2
            total(3) = benchscore(3) * 1.1 + shuttlescore(3) * 1.1 + fortyscore(3) * 1.2 + milescore(3) / 1.2 + vleapscore(3) / 1.1 + handeyescore(3) * 1.3
            total(4) = benchscore(4) / 3 + shuttlescore(4) / 1.25 + fortyscore(4) / 1.25 + milescore(4) / 1.5 + vleapscore(4) * 2 + handeyescore(4) * 2
            total(5) = benchscore(5) + shuttlescore(5) + fortyscore(5) + milescore(5) / 1.5 + vleapscore(5) / 1.5 + handeyescore(5) * 1.25
            total(6) = (benchscore(6) + shuttlescore(6) + fortyscore(6)) / 2 + milescore(6) * 5 + (vleapscore(6) + handeyescore(6)) / 1.5
            total(7) = benchscore(7) / 1.25 + shuttlescore(7) + fortyscore(7) * (2.5) + milescore(7) / 1.25 + vleapscore(7) + handeyescore(7) / 1.25
            total(8) = benchscore(8) / 2 + shuttlescore(8) + fortyscore(8) + milescore(8) + vleapscore(8) + handeyescore(8)
            total(9) = benchscore(9) * 2.5 + (shuttlescore(9) + fortyscore(9) + milescore(9) + vleapscore(9)) / 1.25 + handeyescore(9)
            total(10) = benchscore(10) * 3 + (shuttlescore(10) + fortyscore(10) + milescore(10)) / 3 + vleapscore(10) / 1.5 + handeyescore(10) / 1.25
            total(11) = benchscore(11) * 5 + (shuttlescore(11) + fortyscore(11) + milescore(11) + vleapscore(11) + handeyescore(11)) / 2
            total(12) = benchscore(12) / 2 + (shuttlescore(12) + fortyscore(12) + milescore(12) + vleapscore(12)) / 3 + handeyescore(12) * 3
            total(13) = (benchscore(13) + shuttlescore(13) + fortyscore(13) + milescore(13) + vleapscore(13) + handeyescore(13)) / 1.5
            total(14) = (benchscore(14) + shuttlescore(14) + fortyscore(14) + milescore(14) + vleapscore(14) + handeyescore(14)) / 2
        frmFirst.Hide 'loads second form (main form)
        frmSecond.Show
        For pass = 1 To size - 1 'sorts the arrays from greates value, while also moving the name with the number
            For pos = 1 To 14 - pass
                If total(pos) < total(pos + 1) Then
                temp = total(pos)
                total(pos) = total(pos + 1)
                total(pos + 1) = temp
                tempname = arraysport(pos)
                arraysport(pos) = arraysport(pos + 1)
                arraysport(pos + 1) = tempname
                Else
                End If
            Next pos
        Next pass
        For pos = 1 To 14
            rankname(pos) = arraysport(pos)
        Next pos
        Open App.Path & "\sports.txt" For Input As #1 ' reloads origonal data for future use
        pos = 0
        Do Until EOF(1)
            pos = pos + 1
            Input #1, arraysport(pos), arrayforty(pos), arrayshuttle(pos), arraybench(pos), arraymile(pos), arrayleap(pos), arrayhandeye(pos)
        Loop
        Close #1
    Else
    End If
End Sub

Private Sub cmdCoor_Click()
    Dim Hands, tires, bullseye, hits, Balance As Integer
    Dim total As Single
    total = 0
    Hands = InputBox("How many of the ten balls did you catch?", "Hands") 'prompts user for calulation input
        If Hands >= 10 Then 'compares user input with value to assign score
            total = total + 2
        ElseIf Hands = 9 Then
            total = total + 1.8
        ElseIf Hands = 8 Then
            total = total + 1.6
        ElseIf Hands = 7 Then
            total = total + 1.4
        ElseIf Hands = 6 Then
            total = total + 1.2
        ElseIf Hands = 5 Then
            total = total + 1
        ElseIf Hands = 4 Then
            total = total + 0.8
        ElseIf Hands = 3 Then
            total = total + 0.6
        ElseIf Hands = 2 Then
            total = total + 0.4
        ElseIf Hands = 1 Then
            total = total + 0.2
        ElseIf Hands = 0 Then
            total = total + 0
        End If
    tires = InputBox("How many times did you fall during the tires?", "tires") 'prompts user for calulation input
        If tires >= 10 Then 'compares user input with value to assign score
            total = total + 0
        ElseIf tires = 9 Then
            total = total + 0.2
        ElseIf tires = 8 Then
            total = total + 0.4
        ElseIf tires = 7 Then
            total = total + 0.6
        ElseIf tires = 6 Then
            total = total + 0.8
        ElseIf tires = 5 Then
            total = total + 1
        ElseIf tires = 4 Then
            total = total + 1.2
        ElseIf tires = 3 Then
            total = total + 1.4
        ElseIf tires = 2 Then
            total = total + 1.6
        ElseIf tires = 1 Then
            total = total + 1.8
        ElseIf tires = 0 Then
            total = total + 2
        End If
    bullseye = InputBox("Out of 500 what was your score in the bullseye drill (bulleye=50, next ring=25, next ring=10, ando outer ring=5)", "tires") 'prompts user for calulation input
        If bullseye >= 500 Then 'compares user input with value to assign score
            total = total + 2
        ElseIf bullseye >= 450 Then
            total = total + 1.8
        ElseIf bullseye >= 400 Then
            total = total + 1.6
        ElseIf bullseye >= 350 Then
            total = total + 1.4
        ElseIf bullseye >= 300 Then
            total = total + 1.2
        ElseIf bullseye >= 250 Then
            total = total + 1
        ElseIf bullseye >= 200 Then
            total = total + 0.8
        ElseIf bullseye >= 150 Then
            total = total + 0.6
        ElseIf bullseye >= 100 Then
            total = total + 0.4
        ElseIf bullseye >= 50 Then
            total = total + 0.2
        ElseIf bullseye >= 0 Then
            total = total + 0
        End If
     hits = InputBox("Out of the ten pitches, how many hits did you get in the hitting drill?", "hits") 'prompts user for calulation input
        If hits >= 10 Then 'compares user input with value to assign score
            total = total + 2
        ElseIf hits = 9 Then
            total = total + 1.8
        ElseIf hits = 8 Then
            total = total + 1.6
        ElseIf hits = 7 Then
            total = total + 1.4
        ElseIf hits = 6 Then
            total = total + 1.2
        ElseIf hits = 5 Then
            total = total + 1
        ElseIf hits = 4 Then
            total = total + 0.8
        ElseIf hits = 3 Then
            total = total + 0.6
        ElseIf hits = 2 Then
            total = total + 0.4
        ElseIf hits = 1 Then
            total = total + 0.2
        ElseIf hits = 0 Then
            total = total + 0
        End If
    tires = InputBox("How many times did you fall during balance drill?", "balance") 'prompts user for calulation input
        If Balance > 19 Then 'compares user input with value to assign score
            total = total + 0
        ElseIf Balance > 17 Then
            total = total + 0.2
        ElseIf Balance > 15 Then
            total = total + 0.4
        ElseIf Balance > 13 Then
            total = total + 0.6
        ElseIf Balance > 11 Then
            total = total + 0.8
        ElseIf Balance > 9 Then
            total = total + 1
        ElseIf Balance > 7 Then
            total = total + 1.2
        ElseIf Balance > 5 Then
            total = total + 1.4
        ElseIf Balance > 3 Then
            total = total + 1.6
        ElseIf Balance > 1 Then
            total = total + 1.8
        Else
            total = total + 2
        End If
    MsgBox "Your corrdination rating is " & total, , "rating"
End Sub

Private Sub cmdQuit_Click() 'Quits program
    End
End Sub


Private Sub tmrFour_Timer() 'delays the displaying of text
    PicDisplay.Cls
    PicDisplay.Print "Wait... finish your Calculations"
    PicDisplay.Print "first!"
    tmrFour = False 'stops timer from repeating
End Sub

Private Sub tmrOne_Timer() 'delays the displaying of text
    tmrOne = True
    PicDisplay.Cls
    PicDisplay.Print "You should be outside!"
    tmrOne = False 'stops timer from repeating
End Sub

Private Sub tmrThree_Timer() 'delays the displaying of text
    PicDisplay.Print "No really go outside!"
    tmrThree = False 'stops timer from repeating
End Sub

Private Sub tmrTwo_Timer() 'delays the displaying of text
    PicDisplay.Print "What a great day for sports!"
    tmrTwo = False 'stops timer from repeating
End Sub
