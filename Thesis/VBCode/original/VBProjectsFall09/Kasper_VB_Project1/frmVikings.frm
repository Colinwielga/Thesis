VERSION 5.00
Begin VB.Form frmVikings 
   BackColor       =   &H80000007&
   Caption         =   "Vikings"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRteams 
      Caption         =   "Return to Teams"
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort By Offensive Name"
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      Height          =   4455
      Left            =   5520
      ScaleHeight     =   4395
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmddeffense 
      Caption         =   "Defense"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdOffense 
      Caption         =   "Offense"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5400
      Left            =   360
      Picture         =   "frmVikings.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   8040
   End
End
Attribute VB_Name = "frmVikings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pos As Integer 'dims pos as integer for all buttons
'Author: Brandon Kasper
'Written 10/19/2009
'this form allows the user to click multiple buttons, and view Vikings O and D
'It also sorts the list from an array into order

Private Sub cmddeffense_Click()

    Open App.Path & "\VikingsD.txt" For Input As #1 'opens the file Vikings offense
    Ctr = 0 'sets the value of the counter to 0
    Do Until EOF(1) 'starts the looping and sets it to the end of file
        Ctr = Ctr + 1 'adds a running total
        Input #1, VDnumb(Ctr), VDplayers(Ctr), VDpos(Ctr)
    Loop
    Close #1 'closes the array
     picResults.Cls 'clears the picture box
     picResults.Print "# Name ", "    ", "Position"
    For Pos = 1 To Ctr 'output section
        picResults.Print VDnumb(Pos); VDplayers(Pos); Tab(30); VDpos(Pos)
    Next Pos
End Sub

Private Sub cmdOffense_Click()
    
    Open App.Path & "\VikingsO.txt" For Input As #1 'opens the file Vikings offense
    Ctr = 0 'sets the value of the counter to 0
    Do Until EOF(1) 'starts the looping and sets it to the end of file
        Ctr = Ctr + 1 'adds a running total
        Input #1, VOnumb(Ctr), VOplayers(Ctr), VOpos(Ctr)
    Loop
    Close #1
     picResults.Cls
     picResults.Print "# Name ", "    ", "Position"
    For Pos = 1 To Ctr 'output section
        picResults.Print VOnumb(Pos); VOplayers(Pos); Tab(30); VOpos(Pos)
    Next Pos
End Sub


Private Sub cmdRteams_Click()
    frmVikings.Hide 'hides form from user
    frmTeams.Show   'shows form to the user
End Sub

Private Sub cmdSort_Click()
   Dim pass As Integer, Pos As Integer
   Dim TempOP As String, TempON As Integer, TempOPS As String
   
   Dim I As Integer
   
   For pass = 1 To Ctr - 1 'keeps track of how many passes
    For Pos = 1 To Ctr - pass 'keeps track of how many comparisons
        If VOplayers(Pos) > VOplayers(Pos + 1) Then
            TempOP = VOplayers(Pos) 'exchange values if out of order
            VOplayers(Pos) = VOplayers(Pos + 1)
            VOplayers(Pos + 1) = TempOP
            TempON = VOnumb(Pos) 'matches values with first sorted values
            VOnumb(Pos) = VOnumb(Pos + 1)
            VOnumb(Pos + 1) = TempON
            TempOPS = VOpos(Pos) 'matches values with firts sorted values
            VOpos(Pos) = VOpos(Pos + 1)
            VOpos(Pos + 1) = TempOPS
         End If
     Next Pos
   Next pass
   picResults.Cls 'clears the picture box
   For I = 1 To Ctr 'output section
    picResults.Print VOplayers(I); Tab(20); VOnumb(I); Tab(25); VOpos(I)
   Next I
End Sub


