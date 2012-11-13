VERSION 5.00
Begin VB.Form frmOpponent 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNewopp 
      Caption         =   "Search For A New Opponent"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   4080
      Width           =   7215
   End
   Begin VB.CommandButton cmdBackmenu2 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   0
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lblOpp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Click on the Picture below of the MIAC symbol to find a Saint John's Opponent."
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Image Image5 
      Height          =   1500
      Left            =   4080
      Picture         =   "Opponent.frx":0000
      Top             =   2160
      Width           =   1380
   End
   Begin VB.Image Image4 
      Height          =   1785
      Left            =   7200
      Picture         =   "Opponent.frx":128F
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2520
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   120
      Picture         =   "Opponent.frx":D68C
      Top             =   2880
      Width           =   3120
   End
   Begin VB.Image Image2 
      Height          =   2280
      Left            =   120
      Picture         =   "Opponent.frx":10ECE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2850
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   6120
      Picture         =   "Opponent.frx":199A8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3540
   End
End
Attribute VB_Name = "frmOpponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'2006 MIAC Tennis Tournament Distribution
'Opponent Search Form
'Blake Heymans
'10/27/06
'Opponent Search Form Objective
    'This form is intended to show the user the team match data for the opponent
    'of the user's choice. First it promts the user to enter an opponent.
    'Then the user is promted to click the MIAC symbol with a picture of Minnesota
    'in order to see the results of the team match with the user selected opponent.
    'The data is taken from two files for matches played during the tournament.
'Pictures were taken from Google image search as well as the Saint John's Univerity web site.
    
Private Sub cmdBackmenu2_Click()
    'Brings the user back to Title Form
    frmTitle.Show
    frmOpponent.Hide
End Sub

Private Sub cmdNewopp_Click()
    'Allows the user to start a new opponent search
    Opp = InputBox("Your Choices Are UST, CC, and GAC.                                                                                                                      Be Sure to Capitalize the Letters!", "Enter An Opponent")
    MsgBox "Please Click on the MIAC Symbol of the Image of Minnesota to See The Results of the Match."
End Sub

Private Sub Image5_Click()
Dim O As Integer
Dim Ctr As Integer

    picResults.Cls
        
    'Singles Players Search
    Ctr = 0
    
    Open App.Path & "\Singles.txt" For Input As #1
    
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Playernames(Ctr), Scores(Ctr), Opponent(Ctr), Winloss(Ctr)
    Loop
    
    Close #1
    
    'Prints the titles that will be used
    picResults.Print "Player's Names"; Tab(20); "Scores"; Tab(35); "Win Or A Loss"; Tab(52); "Doubles Partner"
    picResults.Print "==================================================================================================="
    
    'Prints the info that the user is searching for from the file 1
    For O = 1 To Ctr
        If Opp = Opponent(O) Then
            picResults.Print Playernames(O); Tab(20); Scores(O); Tab(35); Winloss(O); Tab(52); "Singles Match"
        End If
    Next O
    
    'Now The Doubles Player Search
    Ctr = 0
    
    Open App.Path & "\Doubles.txt" For Input As #2
    
    Do While Not EOF(2)
        Ctr = Ctr + 1
        Input #2, Playernamesd(Ctr), Scoresd(Ctr), Opponentd(Ctr), Winlossd(Ctr), Dpartner(Ctr)
    Loop
    
    Close #2
    
    'Prints the info that the user is searching for from the file 2
    For O = 1 To Ctr
        If Opp = Opponentd(O) Then
            picResults.Print Playernamesd(O); Tab(20); Scoresd(O); Tab(35); Winlossd(O); Tab(52); Dpartner(O)
        End If
    Next O
End Sub
