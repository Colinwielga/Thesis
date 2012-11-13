VERSION 5.00
Begin VB.Form frmPitch 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   10260
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Johnson_frmPitch.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdPic 
      Caption         =   "Who's Behind the Picture Box???"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "BACK!"
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdPitch 
      Caption         =   "2008 Brewer     Pitching Stats"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.PictureBox picResultsPlayer1 
      Height          =   7215
      Left            =   2280
      ScaleHeight     =   7155
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   1200
      Width           =   11535
   End
End
Attribute VB_Name = "frmPitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Milwaukee Brewers Fan Club Program 2008

'Form Name: Brewer's Pitching

'Author: Matthew Johnson

'Date Written: 11/2/2008

'Objective: In this part of the program, I construct a program that reads pitching
'statistics from a file to educate the user on the quality of Brewer's pitching :-p
'I also make a guessing game within this form that uses inputboxes and the hide and show functions.

Option Explicit
'Here I declare multiple variables into an array that can be used in reading the below
'inputs from a text file under the cmdPitch_Click() button routine. 'These variables
'can be used throughout the form under multiple sections; however, I don't use it in multiple sections.
Dim Player(1 To 40) As String, Team(1 To 40) As String, W(1 To 40) As Integer, L(1 To 40) As Integer
Dim ERA(1 To 40) As Single, G1(1 To 40) As Integer, GS(1 To 40) As Integer, CG(1 To 40) As Integer
Dim SHO(1 To 40) As Integer, SV(1 To 40) As Integer, SVO(1 To 40) As Integer, IP(1 To 40) As Single
Dim ER(1 To 40) As Integer, HR1(1 To 40) As Integer, HBP(1 To 40) As Integer, BB1(1 To 40) As Integer
Dim SO1(1 To 40) As Integer, H1(1 To 40) As Integer, R1(1 To 40) As Integer
Dim ctr As Integer

'Here I allow the user to go back to the initial page.
Private Sub CmdBack_Click()
    picResultsPlayer1.Visible = True
    frmPitch.Hide
    frmIntro.Show

End Sub

'Here is an interesting button.  I have an input box that asks who's behind the picture
'box,and if the user is right, a msgBox tells them they're correct and the pitcher appears
'behind the picture. If they're wrong, a msg box tells them they're wrong.

Private Sub cmdPic_Click()
Dim Sabathia As String, found As Boolean
found = False
Sabathia = InputBox("Who's behind the picture box?", "Which Pitcher is it?") 'setting the input equal to sabathia

    If Sabathia = "C.C. Sabathia" Then 'if input(sabathia) equals "C.C. Sabathia" then the answer is correct
        found = True
        MsgBox "Congratulations, you were correct!!!!", , "Correct Response"
        picResultsPlayer1.Visible = False
    End If
    
    If Not found Then
        MsgBox "You're Wrong!!!! Maybe Next Time!", , "Wrong Response" 'else: msgbox says the user is wrong
    End If

End Sub

'This is a button that reads the pitching stats from a text file. This Do Until Loop reads everything until
'the end of file (EOF), and displays the data within the text file.

Private Sub cmdPitch_Click()

picResultsPlayer1.Visible = True

picResultsPlayer1.Cls
picResultsPlayer1.Print "Player"; Tab(15); "Team"; Tab(23); "W"; Tab(28); "L"; Tab(33); "ERA"; Tab(42); "G"; Tab(47); "GS"; Tab(55); "CG"; Tab(63); "SHO"; Tab(71); "SV"; Tab(79); "SVO"; Tab(87); "IP"; Tab(95); "H"; Tab(103); "R"; Tab(111); "ER"; Tab(118); "HR"; Tab(124); "HBP"; Tab(130); "BB"; Tab(137); "SO"
picResultsPlayer1.Print "************************************************************************************************************************************************************************************************************************************************************************************************************************************"
picResultsPlayer1.Print ""  'printing with double quotes prints exact text

ctr = 0

Open App.Path & "\Pitching Stats.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, Player(ctr), Team(ctr), W(ctr), L(ctr), ERA(ctr), G1(ctr), GS(ctr), CG(ctr), SHO(ctr), SV(ctr), SVO(ctr), IP(ctr), H1(ctr), R1(ctr), ER(ctr), HR1(ctr), HBP(ctr), BB1(ctr), SO1(ctr)
        picResultsPlayer1.Print Player(ctr); Tab(15); Team(ctr); Tab(23); W(ctr); Tab(28); L(ctr); Tab(33); ERA(ctr); Tab(42); G1(ctr); Tab(47); GS(ctr); Tab(55); CG(ctr); Tab(63); SHO(ctr); Tab(71); SV(ctr); Tab(79); SVO(ctr); Tab(87); IP(ctr); Tab(95); H1(ctr); Tab(103); R1(ctr); Tab(111); ER(ctr); Tab(118); HR1(ctr); Tab(124); HBP(ctr); Tab(130); BB1(ctr); Tab(137); SO1(ctr) 'printing the value of a variable (not the name)
    Loop
Close #1
End Sub
