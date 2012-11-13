VERSION 5.00
Begin VB.Form frmAvalanche 
   BackColor       =   &H80000012&
   Caption         =   "Colorado Avalanche"
   ClientHeight    =   13200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmAvalanche.frx":0000
   ScaleHeight     =   13200
   ScaleWidth      =   17250
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdAppearal 
      BackColor       =   &H80000002&
      Caption         =   "Team Appearal "
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   3735
   End
   Begin VB.CommandButton cmdPoints 
      BackColor       =   &H80000002&
      Caption         =   "See Who is Leading the Team in Points"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmdReadFile 
      BackColor       =   &H80000002&
      Caption         =   "View Team Stats"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12375
      Left            =   8520
      ScaleHeight     =   12375
      ScaleWidth      =   6255
      TabIndex        =   1
      Top             =   240
      Width           =   6255
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAvalanche.frx":2F14
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1815
   End
End
Attribute VB_Name = "frmAvalanche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
'The purpose of this form and the other four team forms is to read a team text file and print the team stats as well as compute each players total points and print the results
'Each of the Team forms have a button to view and buy team appearal on a different forms

Private Sub cmdAppearal_Click()
    
    frmAvalanche.Hide
    frmAvalancheStuff.Show
    
End Sub

'Hide Main Menu and show Avalanche Form
Private Sub cmdMainMenu_Click()
frmAvalanche.Hide
frmMainMenu.Show

End Sub


Private Sub cmdPoints_Click()
picResults.Print Tab(6); "Each Players' Total Points"
picResults.Print "_____________________________________"
   picResults.Print Ctr
 For i = 1 To Ctr
    Colpoints(i) = ColGoals(i) + ColAssists(i)
    picResults.Print ColPlayer(i); Tab(20); Colpoints(i)
 Next i
 
End Sub

Private Sub cmdReadFile_Click()



Open App.Path & "\Colorado.txt" For Input As #1
    Ctr = 0

picResults.Print "Player"; Tab(20); "Number"; Tab(30); "Goals"; Tab(40); "Assists"; Tab(50); "+/-"
picResults.Print "****************************************************************"

    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, ColPlayer(Ctr), ColNumber(Ctr), ColGoals(Ctr), ColAssists(Ctr), ColPlusMinus(Ctr)
        picResults.Print ColPlayer(Ctr); Tab(20); ColNumber(Ctr); Tab(30); ColGoals(Ctr); Tab(40); ColAssists(Ctr); Tab(50); ColPlusMinus(Ctr)
        
    Loop
    
Close #1
    picResults.Print
    picResults.Print
    picResults.Print
    
cmdReadFile.Enabled = False
cmdPoints.Enabled = True
End Sub

