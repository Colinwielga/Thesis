VERSION 5.00
Begin VB.Form frmRank 
   BackColor       =   &H00008000&
   Caption         =   "Rank the pitchers by ERA and Strikeouts"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   Picture         =   "frmRank.frx":0000
   ScaleHeight     =   6060
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdERArank 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click to list the pitchers in order of ERA"
      Height          =   1095
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdstrikeoutrank 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click to list the pitchers in order of STRIKEOUTS"
      Height          =   1095
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Go back to the starting diamond"
      Height          =   1095
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.PictureBox picrankings 
      BackColor       =   &H0080FFFF&
      Height          =   4935
      Left            =   3480
      ScaleHeight     =   4875
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MLBPitchers (MLBPitchers.vbp)
'Form Name: frmRank (frmRank.frm)
'Author: Bradley Vrchota
'Date: March 14, 2004
'Purpose: This form shows the user lists of the 20 pitchers ranked
        'lowest(best) to highest ERA with one button and most to
        'least strikeouts thrown with another button

Option Explicit
'dimension the variables used in this form
Dim temppitcher As String, tempERA As Single, tempSO As Integer
Dim PASS As Integer, COMP As Integer

Private Sub cmdBack_Click()
    frmRank.Hide            'take user back to starting form
    frmStart.Show
End Sub

Private Sub cmdERArank_Click()
    'clear picture box
    picrankings.Cls

'use Bubble Sort to arrange names in order of ERA
For PASS = 1 To ctr - 1
    For COMP = 1 To ctr - PASS
        If ERA(COMP) > ERA(COMP + 1) Then
        
            'switch ERA
            tempERA = ERA(COMP)
            ERA(COMP) = ERA(COMP + 1)
            ERA(COMP + 1) = tempERA
            
            'and also pitcher names
            temppitcher = pitcher(COMP)
            pitcher(COMP) = pitcher(COMP + 1)
            pitcher(COMP + 1) = temppitcher
            
        End If
    Next COMP
Next PASS
    
    'print header
    picrankings.Print "Pitcher"; Tab(20); "ERA"
    picrankings.Print "*******************************"
    
    'display pitchers in order of ERA
    For J = 1 To ctr
        picrankings.Print pitcher(J); Tab(20); FormatNumber(ERA(J))
    Next J
    
End Sub


Private Sub cmdstrikeoutrank_Click()
     'clear picture box
    picrankings.Cls

'use Bubble Sort to arrange names in order of strikeouts
For PASS = 1 To ctr - 1
    For COMP = 1 To ctr - PASS
        If strikeouts(COMP) < strikeouts(COMP + 1) Then
        
            'switch strikeouts
            tempSO = strikeouts(COMP)
            strikeouts(COMP) = strikeouts(COMP + 1)
            strikeouts(COMP + 1) = tempSO
            
            'and also pitcher names
            temppitcher = pitcher(COMP)
            pitcher(COMP) = pitcher(COMP + 1)
            pitcher(COMP + 1) = temppitcher
            
        End If
    Next COMP
Next PASS
    
    'print header
    picrankings.Print "Pitcher"; Tab(20); "Strikeouts"
    picrankings.Print "************************************"
    
    'display pitchers in order of strikeouts
    For J = 1 To ctr
        picrankings.Print pitcher(J); Tab(20); strikeouts(J)
    Next J
End Sub
