VERSION 5.00
Begin VB.Form frmChoose 
   Caption         =   "Let's Choose"
   ClientHeight    =   6150
   ClientLeft      =   3750
   ClientTop       =   2940
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   Picture         =   "frmChoose.frx":0000
   ScaleHeight     =   6150
   ScaleWidth      =   8550
   Begin VB.CommandButton cmdPrimaryDem 
      BackColor       =   &H00FF0000&
      Caption         =   "See Which Democratic Cantidate My State Chose in the Primaries"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   4815
   End
   Begin VB.CommandButton cmdFinal 
      Caption         =   "See Whose Views I Agree With Most!"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   3600
      Width           =   4335
   End
   Begin VB.CommandButton cmdPrimary 
      BackColor       =   &H000000FF&
      Caption         =   "See Which Republican Cantidate My State Chose in the Primaries"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   4815
   End
   Begin VB.CommandButton cmdHealth 
      Caption         =   "Health Care"
      Enabled         =   0   'False
      Height          =   975
      Left            =   5400
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnergy 
      Caption         =   "Energry/Environment"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3240
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdTaxes 
      Caption         =   "Tax Policies"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdImmigration 
      Caption         =   "Immigration"
      Enabled         =   0   'False
      Height          =   975
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdIraq 
      Caption         =   "The War in Iraq"
      Height          =   975
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Intro "
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: Let's Choose(frmChoose.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 25, 2008
'PURPOSE:  This form is to present you with five questions to help you decide which cantidate to vote for

Option Explicit

'Takes you back to the intro form
Private Sub cmdBack_Click()
frmIntro.Show
frmChoose.Hide
End Sub

'Takes you to Energy and Environment topics
Private Sub cmdEnergy_Click()

cmdIraq.Enabled = False
cmdImmigration.Enabled = False
cmdTaxes.Enabled = False
cmdEnergy.Enabled = False
cmdHealth.Enabled = True
cmdFinal.Enabled = False

frmEnergy.Show
frmChoose.Hide
End Sub

'Using a bubble sort, this decides who had the most concurrent views and prints the answer
Private Sub cmdFinal_Click()

Dim Pos As Integer
Dim pass As Integer
Dim TempCantidateCtr As Integer
Dim TempCantidate As String
Dim J As Single
Dim Ctr As Integer
Pos = 1
Ctr = 4

For pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - pass
        If CantidateCtr(Pos) > CantidateCtr(Pos + 1) Then
            TempCantidateCtr = CantidateCtr(Pos)
            CantidateCtr(Pos) = CantidateCtr(Pos + 1)
            CantidateCtr(Pos + 1) = TempCantidateCtr
            
            TempCantidate = Cantidate(Pos)
            Cantidate(Pos) = Cantidate(Pos + 1)
            Cantidate(Pos + 1) = TempCantidate
        End If
    Next Pos
Next pass

MsgBox Module1.UsersName & ", The cantidate that your views most agree with is " & Cantidate(4) & "!  Now make sure you get out there and vote!!"

CantidateCtr(1) = 0
CantidateCtr(2) = 0
CantidateCtr(3) = 0
CantidateCtr(4) = 0

cmdIraq.Enabled = True
cmdImmigration.Enabled = False
cmdTaxes.Enabled = False
cmdEnergy.Enabled = False
cmdHealth.Enabled = False
cmdFinal.Enabled = False

End Sub

'Takes you to Healthcare topic
Private Sub cmdHealth_Click()

cmdIraq.Enabled = False
cmdImmigration.Enabled = False
cmdTaxes.Enabled = False
cmdEnergy.Enabled = False
cmdHealth.Enabled = False
cmdFinal.Enabled = True

frmHealth.Show
frmChoose.Hide
End Sub

'Takes you to Immigration topic.
Private Sub cmdImmigration_Click()

cmdIraq.Enabled = False
cmdImmigration.Enabled = False
cmdTaxes.Enabled = True
cmdEnergy.Enabled = False
cmdHealth.Enabled = False
cmdFinal.Enabled = False

frmImmigration.Show
frmChoose.Hide
End Sub

'Takes you to Iraq topic.
Private Sub cmdIraq_Click()

cmdIraq.Enabled = False
cmdImmigration.Enabled = True
cmdTaxes.Enabled = False
cmdEnergy.Enabled = False
cmdHealth.Enabled = False
cmdFinal.Enabled = False

frmIraq.Show
frmChoose.Hide

End Sub

'Promts the user with an input box to enter their state name and displays which Republican cantidate their state chose. Uses a match and stop.
Private Sub cmdPrimary_Click()
Dim Ctr As Integer
Dim YourState As String
Dim States(1 To 50) As String
Dim Winner(1 To 50) As String
Dim Match As Integer
Dim Matched As Boolean
Matched = False

Open App.Path & "\BioTexts\PrimaryResultsRepublican.txt" For Input As #1
Ctr = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, States(Ctr), Winner(Ctr)
Loop

YourState = InputBox("Enter your state", "Your State's Choice for the Primaries")

For Match = 1 To Ctr
    If YourState = States(Match) Then
        MsgBox "Your State's winner was " & Winner(Match)
        Matched = True
    End If
Next Match
If Matched = False Then
    MsgBox "Your state was won by a cantidate that is no longer in the race."
End If

Close #1

End Sub

'Promts the user with an input box to enter their state name and displays which democratic cantidate their state chose. Uses a match and stop
Private Sub cmdPrimaryDem_Click()
Dim Ctr As Integer
Dim YourState As String
Dim States(1 To 50) As String
Dim Winner(1 To 50) As String
Dim Match As Integer
Dim Matched As Boolean
Matched = False


Open App.Path & "\BioTexts\PrimaryResultsDemocratic.txt" For Input As #1
Ctr = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, States(Ctr), Winner(Ctr)
Loop

YourState = InputBox("Enter your state", "Your State's Choice for the Primaries")

For Match = 1 To Ctr
    If YourState = States(Match) Then
        MsgBox "Your State's winner was " & Winner(Match)
        Matched = True
    End If
Next Match
If Matched = False Then
    MsgBox "Your state was won by a cantidate that is no longer in the race."
End If

Close #1

End Sub

'Takes you to Tax policies.
Private Sub cmdTaxes_Click()

cmdIraq.Enabled = False
cmdImmigration.Enabled = False
cmdTaxes.Enabled = False
cmdEnergy.Enabled = True
cmdHealth.Enabled = False
cmdFinal.Enabled = False

frmTaxes.Show
frmChoose.Hide
End Sub
