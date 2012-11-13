VERSION 5.00
Begin VB.Form frmMisc 
   BackColor       =   &H00008000&
   Caption         =   "More information about the pitchers"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   Palette         =   "frmMisc.frx":0000
   Picture         =   "frmMisc.frx":0FE6
   ScaleHeight     =   6000
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWinrank 
      BackColor       =   &H0080FFFF&
      Height          =   5655
      Left            =   5880
      ScaleHeight     =   5595
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click to see how your favorite pitcher would rank in wins"
      Height          =   2175
      Left            =   1920
      Picture         =   "frmMisc.frx":106BD
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox picMisc 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   5160
      Width           =   4215
   End
   Begin VB.CommandButton cmdAverage 
      BackColor       =   &H0000FF00&
      Caption         =   "Find the average ERA and Strikeouts"
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Go back to the starting diamond"
      Height          =   975
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
End
Attribute VB_Name = "frmMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MLBPitchers (MLBPitchers.vbp)
'Form Name: frmMisc (frmMisc.frm)
'Author: Bradley Vrchota
'Date: March 14, 2004
'Purpose: This form has one button that displays the average ERA
        'and stikeouts for the 20 pitchers, another button brings
        'up an input box with allows the user to enter a number of
        'wins and see where that number would rank among those 20
    
Option Explicit

Private Sub cmdAverage_Click()
'dimension the sum variables and average variables
Dim ERAsum As Single, SOsum As Single, ERAave As Single, SOave As Single

    picMisc.Cls             'clear picture box
    
    For J = 1 To ctr            'calculate the sums of the 20 ERA's
        ERAsum = ERAsum + ERA(J)        'and the strikeouts
        SOsum = SOsum + strikeouts(J)
    Next J
    
    'Divide the sums by the counter to get an average ERA and # of strikeouts
    ERAave = ERAsum / ctr
    SOave = SOsum / ctr
    
    'print message and averages
    picMisc.Print "The average ERA of the 20 pitchers is "; FormatNumber(ERAave, 2)
    picMisc.Print "The average strikeouts thrown by the 20 pitchers is"; Round(SOave)
    
End Sub

Private Sub cmdBack_Click()
    frmMisc.Hide            'take user back to start form
    frmStart.Show
End Sub

Private Sub cmdInput_Click()
    Dim N As Integer, rank As Integer
    Dim temppitcher As String, tempwins As Integer
    Dim PASS As Integer, COMP As Integer
    
    'have the user enter a number of wins
    N = InputBox("Enter the number of wins your favorite pitcher had last year")
    
     'clear picture box
    picWinrank.Cls

'use Bubble Sort to arrange names in order of wins so later
'the user's number can be inserted in a ranked place
For PASS = 1 To ctr - 1
    For COMP = 1 To ctr - PASS
        If wins(COMP) < wins(COMP + 1) Then
        
            'switch wins
            tempwins = wins(COMP)
            wins(COMP) = wins(COMP + 1)
            wins(COMP + 1) = tempwins
            
            'and also pitcher names
            temppitcher = pitcher(COMP)
            pitcher(COMP) = pitcher(COMP + 1)
            pitcher(COMP + 1) = temppitcher
            
        End If
    Next COMP
Next PASS
    
    picWinrank.Print "Pitcher"; Tab(25); "Wins"
    picWinrank.Print "************************************"
    
    'set original rank to 1
    rank = 1
    'compare the user's wins to each of the 20 pitchers' wins
    'if user's is less than the pitchers, then add one to the rank
    For J = 1 To ctr
        If N < wins(J) Then
            rank = rank + 1
            picWinrank.Print pitcher(J); Tab(25); wins(J)
        End If
    Next J
    'print the place where the user's pitcher and wins would be
    picWinrank.Print "***Your pitcher***"; Tab(25); N
    
    'compare the user and pitcher wins again to complete the list
    'by displaying those with less wins than the user
    For J = 1 To ctr
        If N >= wins(J) Then
            picWinrank.Print pitcher(J); Tab(25); wins(J)
        End If
    Next J
    
    picWinrank.Print            'separate with a space
    
    'Display appropriate statement depending on where on the list
    'the user's wins would be
    Select Case N
        Case Is >= 35
            picWinrank.Print "Be realistic,"; N; "wins? I doubt it."
        Case Is >= wins(1)
            picWinrank.Print "Your pitcher is an ace!"
        Case Is >= wins(ctr / 2)
            picWinrank.Print "A pretty good pitcher!"
        Case Is >= wins(ctr)
            picWinrank.Print "Well, he made the list!"
        Case Is < wins(ctr)
            picWinrank.Print "Maybe your pitcher should try golf."
    End Select
    
    'leave space and print how many pitchers were ahead to the
    'user's and print the user's wins again
    picWinrank.Print
    picWinrank.Print "There were"; (rank - 1); "pitchers with"
    picWinrank.Print "more than"; N; "wins."
End Sub
