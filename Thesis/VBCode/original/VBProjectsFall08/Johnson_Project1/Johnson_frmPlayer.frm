VERSION 5.00
Begin VB.Form frmPlayer 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   -135
   ClientTop       =   645
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Johnson_frmPlayer.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdHR 
      Caption         =   "The Home Run Hitters"
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search by Player Name to get 2008 Hit Stats"
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "GO BACK"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   8640
      Width           =   2295
   End
   Begin VB.CommandButton cmdAlph 
      Caption         =   "Alphabetize Players in 2008 Hit Stats"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdHit 
      Caption         =   "2008 Hitting Stats (Load Data)"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox PicResultsPlayer 
      Height          =   8175
      Left            =   2880
      ScaleHeight     =   8115
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   240
      Width           =   12375
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Milwaukee Brewers Fan Club Program 2008

'Form Name: 2008 Brewers Hitting Statistics

'Author: Matthew Johnson

'Date Written: 11/2/2008

'Objective: In this part of the program, I construct a series of related
'actions dealing with showing the 2008 Brewers Hitting Statistics in different ways
'using my VB skills. It demonstrates my abilities to bubble sort, to exhaustive search,
'and work with multiple forms...

Option Explicit
'Here I declare variables that I use throughout several button commands in the form.
Dim Player(1 To 40) As String, Team(1 To 40) As String, pos(1 To 40) As String, G(1 To 40) As Integer, AB(1 To 40) As Integer, R(1 To 40) As Integer, H(1 To 40) As Integer, second(1 To 40) As Integer, three(1 To 40) As Integer, HR(1 To 40) As Integer, RBI(1 To 40) As Integer, TB(1 To 40) As Integer, BB(1 To 40) As Integer, SO(1 To 40) As Integer, SB(1 To 40) As Integer, CS(1 To 40) As Integer, OBP(1 To 40) As Single, SLG(1 To 40) As Single, AVG(1 To 40) As Single
Dim ctr As Integer

'Here I just enter a file input.  I open a text file and print it out.  It shows the hitting
'stats of the Milwaukee Brewers....
Private Sub cmdHit_Click()
PicResultsPlayer.Cls
PicResultsPlayer.Print "Player"; Tab(15); "Team"; Tab(23); "POS"; Tab(31); "G"; Tab(39); "AB"; Tab(47); "R"; Tab(55); "H"; Tab(63); "2B"; Tab(71); "3B"; Tab(79); "HR"; Tab(87); "RBI"; Tab(95); "TB"; Tab(103); "BB"; Tab(111); "SO"; Tab(118); "SB"; Tab(124); "CS"; Tab(130); "OBP"; Tab(137); "SLG"; Tab(145); "AVG"
PicResultsPlayer.Print "*****************************************************************************************************************************************************************************************************************************************************************************************************************************************"
PicResultsPlayer.Print "" 'printing with double quotes prints exact text

ctr = 0

Open App.Path & "\Hitting Stats.txt" For Input As #1 'Open text document
    Do Until EOF(1) 'loop through the file until the end
        ctr = ctr + 1
        Input #1, Player(ctr), Team(ctr), pos(ctr), G(ctr), AB(ctr), R(ctr), H(ctr), second(ctr), three(ctr), HR(ctr), RBI(ctr), TB(ctr), BB(ctr), SO(ctr), SB(ctr), CS(ctr), OBP(ctr), SLG(ctr), AVG(ctr)
        PicResultsPlayer.Print Player(ctr); Tab(15); Team(ctr); Tab(23); pos(ctr); Tab(31); G(ctr); Tab(39); AB(ctr); Tab(47); R(ctr); Tab(55); H(ctr); Tab(63); second(ctr); Tab(71); three(ctr); Tab(79); HR(ctr); Tab(87); RBI(ctr); Tab(95); TB(ctr); Tab(103); BB(ctr); Tab(111); SO(ctr); Tab(118); SB(ctr); Tab(124); CS(ctr); Tab(130); OBP(ctr); Tab(137); SLG(ctr); Tab(145); AVG(ctr)
    Loop
Close #1
End Sub
'Here I do a bubble sort making the aforementioned information be put in alphabetical order.
Private Sub cmdAlph_Click()
Dim Pass As Integer, pos1 As Integer, A As Integer, tempPlayer As String, tempteam As String, temppos As String, tempG As Integer, TempAB As Integer, TempR As Integer
Dim tempH As Integer, tempsecond As Integer, tempThree As Integer, tempHR As Integer, TempRBI As Integer, TempTB As Integer, TempBB As Integer, TempSO As Integer
Dim TempSB As Integer, TempCS As Integer, TempOBP As Single, TempSLG As Single, TempAVG As Single

PicResultsPlayer.Cls
PicResultsPlayer.Print "Player"; Tab(15); "Team"; Tab(23); "POS"; Tab(31); "G"; Tab(39); "AB"; Tab(47); "R"; Tab(55); "H"; Tab(63); "2B"; Tab(71); "3B"; Tab(79); "HR"; Tab(87); "RBI"; Tab(95); "TB"; Tab(103); "BB"; Tab(111); "SO"; Tab(118); "SB"; Tab(124); "CS"; Tab(130); "OBP"; Tab(137); "SLG"; Tab(145); "AVG"
PicResultsPlayer.Print "*****************************************************************************************************************************************************************************************************************************************************************************************************************************************"
PicResultsPlayer.Print "" ' printing with double quotes prints exact text
ctr = 25
For Pass = 1 To ctr - 1
For pos1 = 1 To ctr - Pass
     If (Player(pos1) > Player(pos1 + 1)) Then
        tempteam = Team(pos1)
        Team(pos1) = Team(pos1 + 1)
        Team(pos1 + 1) = tempteam
        tempPlayer = Player(pos1)
        Player(pos1) = Player(pos1 + 1)
        Player(pos1 + 1) = tempPlayer
        temppos = pos(pos1)
        pos(pos1) = pos(pos1 + 1)
        pos(pos1 + 1) = temppos
        tempG = G(pos1)
        G(pos1) = G(pos1 + 1)
        G(pos1 + 1) = tempG
        TempAB = AB(pos1)
        AB(pos1) = AB(pos1 + 1)
        AB(pos1 + 1) = TempAB
        TempR = R(pos1)
        R(pos1) = R(pos1 + 1)
        R(pos1 + 1) = TempR
        tempH = H(pos1)
        H(pos1) = H(pos1 + 1)
        H(pos1 + 1) = tempH
        tempsecond = second(pos1)
        second(pos1) = second(pos1 + 1)
        second(pos1 + 1) = tempsecond
        tempThree = three(pos1)
        three(pos1) = three(pos1 + 1)
        three(pos1 + 1) = tempThree
        tempHR = HR(pos1)
        HR(pos1) = HR(pos1 + 1)
        HR(pos1 + 1) = tempHR
        TempRBI = RBI(pos1)
        RBI(pos1) = RBI(pos1 + 1)
        RBI(pos1 + 1) = TempRBI
        TempTB = TB(pos1)
        TB(pos1) = TB(pos1 + 1)
        TB(pos1 + 1) = TempTB
        TempBB = BB(pos1)
        BB(pos1) = BB(pos1 + 1)
        BB(pos1 + 1) = TempBB
        TempSO = SO(pos1)
        SO(pos1) = SO(pos1 + 1)
        SO(pos1 + 1) = TempSO
        TempSB = SB(pos1)
        SB(pos1) = SB(pos1 + 1)
        SB(pos1 + 1) = TempSB
        TempCS = CS(pos1)
        CS(pos1) = CS(pos1 + 1)
        CS(pos1 + 1) = TempCS
        TempOBP = OBP(pos1)
        OBP(pos1) = OBP(pos1 + 1)
        OBP(pos1 + 1) = TempOBP
        TempSLG = SLG(pos1)
        SLG(pos1) = SLG(pos1 + 1)
        SLG(pos1 + 1) = TempSLG
        TempAVG = AVG(pos1)
        AVG(pos1) = AVG(pos1 + 1)
        AVG(pos1 + 1) = TempAVG
    End If
Next pos1
Next Pass

For A = 1 To ctr
  PicResultsPlayer.Print Player(A); Tab(15); Team(A); Tab(23); pos(A); Tab(31); G(A); Tab(39); AB(A); Tab(47); R(A); Tab(55); H(A); Tab(63); second(A); Tab(71); three(A); Tab(79); HR(A); Tab(87); RBI(A); Tab(95); TB(A); Tab(103); BB(A); Tab(111); SO(A); Tab(118); SB(A); Tab(124); CS(A); Tab(130); OBP(A); Tab(137); SLG(A); Tab(145); AVG(A) ' printing the value of a variable (not the name)
Next A

End Sub
'Here, the user can go back to the initial page
Private Sub CmdBack_Click()
    frmIntro.Show
    frmPlayer.Hide
End Sub

'This is a Match/Stop search, where the user can find a specific player's stats.
Private Sub cmdSearch_Click()
Dim searchPlayer As String, found As Boolean

    PicResultsPlayer.Cls
    PicResultsPlayer.Print "Player"; Tab(15); "Team"; Tab(23); "POS"; Tab(31); "G"; Tab(39); "AB"; Tab(47); "R"; Tab(55); "H"; Tab(63); "2B"; Tab(71); "3B"; Tab(79); "HR"; Tab(87); "RBI"; Tab(95); "TB"; Tab(103); "BB"; Tab(111); "SO"; Tab(118); "SB"; Tab(124); "CS"; Tab(130); "OBP"; Tab(137); "SLG"; Tab(145); "AVG"
    PicResultsPlayer.Print "*****************************************************************************************************************************************************************************************************************************************************************************************************************************************"
    PicResultsPlayer.Print "" ' printing with double quotes prints exact text

found = False
searchPlayer = InputBox("Enter a name of a brewer with the first initial of the first name followed by the last name (Example: R Braun): ", "Enter Player Name") 'Setting a variable equal to what inputed in the input vox
ctr = 1

Do While ((Not found) And (ctr <= 40))
    If Player(ctr) = searchPlayer Then
        found = True 'this tells the program that the exhaustive search found someone that met the criteria of the search.
        PicResultsPlayer.Print Player(ctr); Tab(15); Team(ctr); Tab(23); pos(ctr); Tab(31); G(ctr); Tab(39); AB(ctr); Tab(47); R(ctr); Tab(55); H(ctr); Tab(63); second(ctr); Tab(71); three(ctr); Tab(79); HR(ctr); Tab(87); RBI(ctr); Tab(95); TB(ctr); Tab(103); BB(ctr); Tab(111); SO(ctr); Tab(118); SB(ctr); Tab(124); CS(ctr); Tab(130); OBP(ctr); Tab(137); SLG(ctr); Tab(145); AVG(ctr) ' printing the value of a variable (not the name)
    End If
        ctr = ctr + 1
    Loop
    
End Sub

'Here I do an exhaustive search, where the user can find a player with homers equal or more than
'the entered number
Private Sub cmdHR_Click()
Dim enterHR As Integer, found As Boolean, N As Integer
PicResultsPlayer.Cls

found = False

enterHR = InputBox("Find Players with homers equal or more than the entered number: ", "Enter an Amount") 'Setting a variable equal to what inputed in the input box

    For N = 1 To ctr
        If HR(N) >= enterHR Then 'If what was inputed in the input box is less the the number of homers a 2008 brewer hit, the program will print their name
            PicResultsPlayer.Print Player(N) 'printing the value of a variable (not the name)
            found = True 'this tells the program that the exhaustive search found someone that met the criteria of the search.
        End If
    Next N

    If found = False Then
        PicResultsPlayer.Print "There are no players with homers equal or more than the entered number." ' printing with double quotes prints exact text
    End If
    
End Sub
