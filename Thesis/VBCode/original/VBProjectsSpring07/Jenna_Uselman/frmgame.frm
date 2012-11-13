VERSION 5.00
Begin VB.Form frmgame 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Interactive Game"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.OptionButton butA 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   7
      Top             =   4800
      Width           =   4335
   End
   Begin VB.CommandButton cmdmenu 
      BackColor       =   &H80000009&
      Caption         =   "return to menu"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdnextqs 
      BackColor       =   &H80000009&
      Caption         =   "start!"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   3015
   End
   Begin VB.OptionButton butD 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   6480
      Width           =   4575
   End
   Begin VB.OptionButton butC 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Top             =   5880
      Width           =   4335
   End
   Begin VB.OptionButton butB 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   5280
      Width           =   4575
   End
   Begin VB.Label lbluser 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      TabIndex        =   8
      Top             =   4920
      Width           =   3975
   End
   Begin VB.Image imgcover 
      Height          =   2655
      Left            =   840
      Picture         =   "frmgame.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   8580
   End
   Begin VB.Image imgsix 
      Height          =   2475
      Left            =   3240
      Picture         =   "frmgame.frx":0E72
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   4380
   End
   Begin VB.Image imgoct 
      Height          =   2610
      Left            =   3960
      Picture         =   "frmgame.frx":700F
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2610
   End
   Begin VB.Image imggrey 
      Height          =   2640
      Left            =   2640
      Picture         =   "frmgame.frx":97B4
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   4950
   End
   Begin VB.Image imgbetty 
      Height          =   2250
      Left            =   3480
      Picture         =   "frmgame.frx":117C6
      Top             =   1320
      Width           =   3270
   End
   Begin VB.Image imglost 
      Height          =   2640
      Left            =   3000
      Picture         =   "frmgame.frx":1FA67
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3960
   End
   Begin VB.Image imgemerg 
      Height          =   2400
      Left            =   1680
      Picture         =   "frmgame.frx":240E7
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   7710
   End
   Begin VB.Image imgjim 
      Height          =   2640
      Left            =   3120
      Picture         =   "frmgame.frx":2AF8D
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   4140
   End
   Begin VB.Image imggeorge 
      Height          =   1575
      Left            =   1800
      Picture         =   "frmgame.frx":2FE28
      Top             =   1560
      Width           =   6375
   End
   Begin VB.Image imgboston 
      Height          =   2520
      Left            =   3480
      Picture         =   "frmgame.frx":33685
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Image imgbrian 
      Height          =   2250
      Left            =   3600
      Picture         =   "frmgame.frx":3B308
      Top             =   1200
      Width           =   3270
   End
   Begin VB.Image Imgdanc 
      Height          =   2640
      Left            =   3840
      Picture         =   "frmgame.frx":4B168
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Image imgDH 
      Height          =   2745
      Left            =   2160
      Picture         =   "frmgame.frx":595BC
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Image imgextreme 
      Height          =   2490
      Left            =   2880
      Picture         =   "frmgame.frx":608ED
      Top             =   1200
      Width           =   4035
   End
   Begin VB.Image imgAFV 
      Height          =   2400
      Left            =   3720
      Picture         =   "frmgame.frx":6408F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Image imgbrothers 
      Height          =   2400
      Left            =   3000
      Picture         =   "frmgame.frx":6A0EB
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3840
   End
   Begin VB.Label lblshowname 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label lblquestion 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   3720
      Width           =   10095
   End
End
Attribute VB_Name = "frmgame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ans As String
Dim guess As String
Dim wrong As Integer
Dim correct As Integer
Dim Ctr As Integer

'Purpose of Form: This form allows the user to play the trivia game. The game consists of 15
'                 trivia questions. The user selects their answer and then clicks the next question button.
'                 After the user finishes the trivia questions. Clicking on the results button, displays a
'                 message box that shows how many correct and rates how well they performed.

Private Sub butA_Click() 'This button adds one to correct if "A" is selected as the answer and "A" is correct.
                         'IF statement copied from Pradeep de Noronha's "Millionaire Project".
   guess = "A"

        If guess = ans Then
            correct = correct + 1
        End If
    Call restart
        
End Sub

Private Sub butB_Click() 'This button adds one to correct if "B" is selected as the answer and "B" is correct.
                         'IF statement copied from Pradeep de Noronha's "Millionaire Project".
    guess = "B"
    
        If guess = ans Then
            correct = correct + 1
        End If
    Call restart
        
End Sub

Private Sub butC_Click() 'This button adds one to correct if "C" is selected as the answer and "C" is correct.
                         'IF statement copied from Pradeep de Noronha's "Millionaire Project".
    guess = "C"
    
        If guess = ans Then
            correct = correct + 1
        End If
    Call restart
        
End Sub

Private Sub butD_Click() 'This button adds one to correct if "D" is selected as the answer and "D" is correct.
                         'IF statement copied from Pradeep de Noronha's "Millionaire Project".
    guess = "D"
    
        If guess = ans Then
            correct = correct + 1
        End If
    Call restart
    
        
End Sub

Private Sub cmdmenu_Click() 'This command button allows the user to return to the menu form.
    frmMain.Show
    frmgame.Hide
End Sub

Private Sub cmdnextqs_Click() 'This command button runs the trivia game.
                              'The IF ctr is copied from Pradeep de Noronha's "Millionaire Project".


lbluser.Caption = "Go " & uname & "!"
cmdnextqs.Caption = "Next Question"
cmdnextqs.BackColor = &HFFFFFF 'Changes the color of the "Next Question" command button
cmdnextqs.Enabled = True
Ctr = Ctr + 1
lblshowname.Visible = True
lblquestion.Visible = True
butA.Visible = True
butB.Visible = True
butC.Visible = True
butD.Visible = True
butA.Enabled = True
butB.Enabled = True
butC.Enabled = True
butD.Enabled = True

imgAFV.Visible = False 'The following images appear when the questions related to the picture is asked.
imgextreme.Visible = False
imgDH.Visible = False
imgbrothers.Visible = False
Imgdanc.Visible = False
imgbrian.Visible = False
imgboston.Visible = False
imggeorge.Visible = False
imgjim.Visible = False
imgemerg.Visible = False
imglost.Visible = False
imgbetty.Visible = False
imggrey.Visible = False
imgoct.Visible = False
imgsix.Visible = False
imgcover.Visible = True

If Ctr = 1 Then
    lblshowname.Caption = "1. America's Funniest Home Videos"
    cmdnextqs.BackColor = &H808080
    imgAFV.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "What other television show does Tom Bergeron host?"
    butA.Caption = "A: American Idol(FOX)"
    butB.Caption = "B: Identity (NBC)"
    butC.Caption = "C: Dancing with the Stars (ABC)"
    butD.Caption = "D: The Amazing Race (CBS)"
    ans = "C"
    
ElseIf Ctr = 2 Then
    lblshowname.Caption = "2. Extreme Makeover: Home Edition"
    cmdnextqs.BackColor = &H80FF&
    lblquestion.Caption = "Who is the team leader of Extreme Makeover: Home Edition?"
    imgextreme.Visible = True
    imgcover.Visible = False
    butA.Caption = "A: Michael Moloney"
    butB.Caption = "B: Ty Pennington"
    butC.Caption = "C: Paul DiMeo"
    butD.Caption = "D: Paige Hemmis"
    ans = "B"
    
ElseIf Ctr = 3 Then
    lblshowname.Caption = "3. Desperate Housewives"
    cmdnextqs.BackColor = &HFF&
    imgDH.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "Which actress plays the role of Lynette Scavo in Desperate Housewives?"
    butA.Caption = "A: Felicity Hoffman"
    butB.Caption = "B: Andrea Bowen"
    butC.Caption = "C: Marcia Cross"
    butD.Caption = "D: Eva Longoria"
    ans = "A"

ElseIf Ctr = 4 Then
    lblshowname.Caption = "4. Brothers and Sisters"
    cmdnextqs.BackColor = &HC000&
    imgbrothers.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "In the television series, Brothers and Sisters, which family member dies?"
    butA.Caption = "A: The Mother"
    butB.Caption = "B: A Sister"
    butC.Caption = "C: The Father"
    butD.Caption = "D: A Brother"
    ans = "C"

ElseIf Ctr = 5 Then
    lblshowname.Caption = "5. Dancing with the Stars"
    cmdnextqs.BackColor = &HFF0000
    Imgdanc.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "Who is Cheryl Burke(two-time winner)'s  new partner in the fourth season?"
    butA.Caption = "A: Billy Ray Cyrus"
    butB.Caption = "B: Ian Ziering"
    butC.Caption = "C: Joey Fatone"
    butD.Caption = "D: John Ratzenberger"
    ans = "B"
    
ElseIf Ctr = 6 Then
    lblshowname.Caption = "6. What About Brian"
    cmdnextqs.BackColor = &H80C0FF
    imgbrian.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "What character did Brian have feelings for in the 1st season?"
    butA.Caption = "A: Marjorie"
    butB.Caption = "B: Deena"
    butC.Caption = "C: Ivy"
    butD.Caption = "D: Natasha"
    ans = "A"
    
ElseIf Ctr = 7 Then
    lblshowname.Caption = "7. Boston Legal"
    cmdnextqs.BackColor = &HC0C0&
    imgboston.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "In Boston Legal, what is the name of the law firm?"
    butA.Caption = "A: Crane, Shore and Chase"
    butB.Caption = "B: Bauer, Schmidt and Shore"
    butC.Caption = "C: Crane, Poole and Schmidt"
    butD.Caption = "D: Bell, Chase and Shore"
    ans = "C"
    
ElseIf Ctr = 8 Then
    lblshowname.Caption = "8. George Lopez"
    cmdnextqs.BackColor = &HFFFFFF
    imggeorge.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "How many seasons has this show been on the air?"
    butA.Caption = "A: Three Seasons"
    butB.Caption = "B: Six Seasons"
    butC.Caption = "C: Five Seasons"
    butD.Caption = "D: Eight Seasons"
    ans = "B"
    
ElseIf Ctr = 9 Then
    lblshowname.Caption = "9. According to Jim"
    cmdnextqs.BackColor = &H80C0FF
    imgjim.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "Kimberly Williams-Paisley plays the role of Jim's _______."
    butA.Caption = "A: sister-in-law"
    butB.Caption = "B: friend's wife"
    butC.Caption = "C: sister"
    butD.Caption = "D: cousin"
    ans = "A"
    
ElseIf Ctr = 10 Then
    lblshowname.Caption = "10. In Case of Emergency"
    cmdnextqs.BackColor = &H80FF&
    imgemerg.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "Actress Lori Loughlin plays the role of Joanna. What other comedy series did she star in?"
    butA.Caption = "A: Step by Step"
    butB.Caption = "B: Boy Meets World"
    butC.Caption = "C: Full House"
    butD.Caption = "D: Who's the Boss?"
    ans = "C"

ElseIf Ctr = 11 Then
    lblshowname.Caption = "11. Lost"
    cmdnextqs.BackColor = &H808000
    imglost.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "Claire gives birth to a baby boy. What did she name the baby?"
    butA.Caption = "A: Josh"
    butB.Caption = "B: Aaron"
    butC.Caption = "C: Charlie"
    butD.Caption = "D: Matthew"
    ans = "B"

ElseIf Ctr = 12 Then
    lblshowname.Caption = "12. Ugly Betty"
    cmdnextqs.BackColor = &HFFC0FF
    imgbetty.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "This famous Latino actress is the executive producer of Ugly Betty."
    butA.Caption = "A: Salma Hayek"
    butB.Caption = "B: Penelope Cruz"
    butC.Caption = "C: Jennifer Lopez"
    butD.Caption = "D: Eva Longoria"
    ans = "A"

ElseIf Ctr = 13 Then
    lblshowname.Caption = "13. Grey's Anatomy"
    cmdnextqs.BackColor = &HFFFFC0
    imggrey.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "Izzie becomes engaged at the end of season 2 to a patient. What is his name?"
    butA.Caption = "A: David"
    butB.Caption = "B: Daniel"
    butC.Caption = "C: Denny"
    butD.Caption = "D: Derek"
    ans = "C"

ElseIf Ctr = 14 Then
    lblshowname.Caption = "14. October Road"
    cmdnextqs.BackColor = &H40C0&
    imgoct.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "Hannah is portrayed by the same actress who played Donna on That 70's Show. Her name is ______."
    butA.Caption = "A: Laura Prepon"
    butB.Caption = "B: Mila Kunis"
    butC.Caption = "C: Debra Jo Rupp"
    butD.Caption = "D: Odette Yustman"
    ans = "A"

ElseIf Ctr = 15 Then
    lblshowname.Caption = "15. Six Degrees"
    cmdnextqs.BackColor = &HFFFF&
    imgsix.Visible = True
    imgcover.Visible = False
    lblquestion.Caption = "This show involves the belief that everyone is connected. What big city is the setting of this show?"
    butA.Caption = "A: Los Angeles"
    butB.Caption = "B: New York City"
    butC.Caption = "C: Boston"
    butD.Caption = "D: Denver"
    ans = "B"
    
ElseIf Ctr = 16 Then
    cmdnextqs.Caption = "Results"
    imgcover.Visible = False

ElseIf Ctr = 17 Then 'Displays the results of the trivia questions.
    Select Case correct
        Case 0 To 3
            MsgBox uname & ", your results are: " & correct & "/15. Just in case you don't know: ABC is on channel 5!"
        Case 4 To 7
            MsgBox uname & ", your results are: " & correct & "/15. Not bad. Try adding ABC to your daily routine!"
        Case 8 To 10
            MsgBox uname & ", your results are:" & correct & "/15. Good. At least you watch a few of ABC's hit shows!"
        Case 11 To 60
            MsgBox uname & ", your results are:" & correct & "/15. Great job! You are the ultimate ABC fan!"
    End Select
    

   
    cmdnextqs.Caption = "Play again?"
    Ctr = 0
    correct = 0
   
    
    
End If


    

    

End Sub

Private Sub restart() 'This subaction makes the choices disappear when a answer is selected.
                      'This subaction is based from Pradeep de Noronha's "Millionaire Project".
    cmdnextqs.Enabled = True
    butA.Visible = False
    butB.Visible = False
    butC.Visible = False
    butD.Visible = False
End Sub



