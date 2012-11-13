VERSION 5.00
Begin VB.Form frmpitching 
   BackColor       =   &H00800000&
   Caption         =   "Pitching Statistics--2005 Houston Astros"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form3"
   ScaleHeight     =   7530
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      BackColor       =   &H000000FF&
      Height          =   1335
      Left            =   1800
      ScaleHeight     =   1275
      ScaleWidth      =   2955
      TabIndex        =   12
      Top             =   5760
      Width           =   3015
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H000000FF&
      Height          =   1335
      Left            =   480
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdoffense 
      Caption         =   "View Offensive Statistics"
      Height          =   735
      Left            =   3720
      TabIndex        =   9
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   6600
      TabIndex        =   8
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   8280
      TabIndex        =   7
      Top             =   6360
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000FF&
      FillColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   3480
      ScaleHeight     =   3915
      ScaleWidth      =   5595
      TabIndex        =   6
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton cmdlosses 
      Caption         =   "Sort Statistics by Fewest Losses (L)"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton cmdinnings 
      Caption         =   "Sort Statistics by Most Innings Pitched (IP)"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdstrikeouts 
      Caption         =   "Sort Statistics by Most Strikeouts (K)"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton cmdera 
      Caption         =   "Sort Statistics by Lowest Earned Run Average (ERA)"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton cmdwins 
      Caption         =   "Sort Statistics by Most Wins (W)"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load 2005 Houston Astros  Pitching Statistics"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lbltom 
      BackColor       =   &H000000FF&
      Caption         =   "Tom Wentzell"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   1095
   End
End
Attribute VB_Name = "frmpitching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2005 Houston Astros Statistics(Wentzell_Tom_Project)
'frmpitching (frmpitching.frm)
'Tom Wentzell
'October 30, 2005
'The purpose of this form is to display individual pitching statistics for each player
'of the Houston Astros who recorded at least 50 innings pitched in 2005.  It can sort
'the players by production in different statistical categories, which allows the user to
'view the leaders of each category.

'Declare form level variables.  All of these variables will be used in multiple subroutines.
Option Explicit
Dim Player(1 To 20) As String, Wins(1 To 20) As Integer, Losses(1 To 20) As Integer
Dim Innings(1 To 20) As Double, Strikeouts(1 To 20) As Integer, ERA(1 To 20) As Double
Dim CTR As Integer, Pass As Integer, Comp As Integer, J As Integer, tempname As String

'This button loads offensive data from a data file into the program in the form of an
'array.  The data is given for 10 Astros pitchers and it covers five statistical categories.
'This data will be displayed in its original form along with headings displaying
'the information  given in each column.
Private Sub cmdload_Click()
picResults.Cls
picResults.Print "Player"; Tab(25); "W"; Tab(32); "L"; Tab(39); "IP"; Tab(49); "K"; Tab(58); "ERA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(39); "-----"; Tab(49); "-----"; Tab(58); "-----"
picResults.Print ""
Open App.Path & "\Pitching Stats.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Player(CTR), Wins(CTR), Losses(CTR), Innings(CTR), Strikeouts(CTR), ERA(CTR)
    picResults.Print Player(CTR); Tab(25); Wins(CTR); Tab(32); Losses(CTR); Tab(39); Innings(CTR); Tab(49); Strikeouts(CTR); Tab(58); ERA(CTR)
Loop
Close
End Sub

'This button uses a bubble sort to rearrange the players according to highest number of
'wins recored.  It then reprints the data in this new sequence.  It also displays the
'picture of the player who lead the team in this statistical category as well as the
'amount of wins he recorded in leading the team.
Private Sub cmdwins_Click()
picResults.Cls
picResults.Print "Player"; Tab(25); "W"; Tab(32); "L"; Tab(39); "IP"; Tab(49); "K"; Tab(58); "ERA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(39); "-----"; Tab(49); "-----"; Tab(58); "-----"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If Wins(Comp) < Wins(Comp + 1) Then
            tempname = Wins(Comp)
            Wins(Comp) = Wins(Comp + 1)
            Wins(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Losses(Comp)
            Losses(Comp) = Losses(Comp + 1)
            Losses(Comp + 1) = tempname
            tempname = Innings(Comp)
            Innings(Comp) = Innings(Comp + 1)
            Innings(Comp + 1) = tempname
            tempname = Strikeouts(Comp)
            Strikeouts(Comp) = Strikeouts(Comp + 1)
            Strikeouts(Comp + 1) = tempname
            tempname = ERA(Comp)
            ERA(Comp) = ERA(Comp + 1)
            ERA(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Wins(J); Tab(32); Losses(J); Tab(39); Innings(J); Tab(49); Strikeouts(J); Tab(58); ERA(J)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Oswalt.jpg")
picResults2.Cls
picResults2.Print "Roy Oswalt lead the 2005 Houston"
picResults2.Print "Astros in total wins with 20."
End Sub

'This button uses a bubble sort to rearrange the players according to lowest number of
'losses recorded.  It then reprints the data in this new sequence.  It also displays the
'picture of the player who recorded the fewest losses as well as the amount of losses
'he recorded.
Private Sub cmdlosses_Click()
picResults.Cls
picResults.Print "Player"; Tab(25); "W"; Tab(32); "L"; Tab(39); "IP"; Tab(49); "K"; Tab(58); "ERA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(39); "-----"; Tab(49); "-----"; Tab(58); "-----"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If Losses(Comp) > Losses(Comp + 1) Then
            tempname = Losses(Comp)
            Losses(Comp) = Losses(Comp + 1)
            Losses(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Wins(Comp)
            Wins(Comp) = Wins(Comp + 1)
            Wins(Comp + 1) = tempname
            tempname = Innings(Comp)
            Innings(Comp) = Innings(Comp + 1)
            Innings(Comp + 1) = tempname
            tempname = Strikeouts(Comp)
            Strikeouts(Comp) = Strikeouts(Comp + 1)
            Strikeouts(Comp + 1) = tempname
            tempname = ERA(Comp)
            ERA(Comp) = ERA(Comp + 1)
            ERA(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Wins(J); Tab(32); Losses(J); Tab(39); Innings(J); Tab(49); Strikeouts(J); Tab(58); ERA(J)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Wheeler.jpg")
picResults2.Cls
picResults2.Print "Dan Wheeler lead the 2005 Houston"
picResults2.Print "Astros in fewest losses with 3."
End Sub

'This button uses a bubble sort to rearrange the players according to highest number of
'innings pitched.  It then reprints the data in this new sequence.  It also displays the
'picture of the player who lead the team in this statistical category as well as the
'amount of innings he pitched in leading the team.
Private Sub cmdinnings_Click()
picResults.Cls
picResults.Print "Player"; Tab(25); "W"; Tab(32); "L"; Tab(39); "IP"; Tab(49); "K"; Tab(58); "ERA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(39); "-----"; Tab(49); "-----"; Tab(58); "-----"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If Innings(Comp) < Innings(Comp + 1) Then
            tempname = Innings(Comp)
            Innings(Comp) = Innings(Comp + 1)
            Innings(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Losses(Comp)
            Losses(Comp) = Losses(Comp + 1)
            Losses(Comp + 1) = tempname
            tempname = Wins(Comp)
            Wins(Comp) = Wins(Comp + 1)
            Wins(Comp + 1) = tempname
            tempname = Strikeouts(Comp)
            Strikeouts(Comp) = Strikeouts(Comp + 1)
            Strikeouts(Comp + 1) = tempname
            tempname = ERA(Comp)
            ERA(Comp) = ERA(Comp + 1)
            ERA(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Wins(J); Tab(32); Losses(J); Tab(39); Innings(J); Tab(49); Strikeouts(J); Tab(58); ERA(J)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Oswalt.jpg")
picResults2.Cls
picResults2.Print "Roy Oswalt lead the 2005 Houston"
picResults2.Print "Astros in innings pitched with 241.2."
End Sub

'This button uses a bubble sort to rearrange the players according to highest number of
'strikeouts recorded.  It then reprints the data in this new sequence.  It also displays
'the picture of the player who lead the team in this statistical category as well as the
'amount of strikeouts he recorded in leading the team.
Private Sub cmdstrikeouts_Click()
picResults.Cls
picResults.Print "Player"; Tab(25); "W"; Tab(32); "L"; Tab(39); "IP"; Tab(49); "K"; Tab(58); "ERA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(39); "-----"; Tab(49); "-----"; Tab(58); "-----"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If Strikeouts(Comp) < Strikeouts(Comp + 1) Then
            tempname = Strikeouts(Comp)
            Strikeouts(Comp) = Strikeouts(Comp + 1)
            Strikeouts(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Losses(Comp)
            Losses(Comp) = Losses(Comp + 1)
            Losses(Comp + 1) = tempname
            tempname = Wins(Comp)
            Wins(Comp) = Wins(Comp + 1)
            Wins(Comp + 1) = tempname
            tempname = Innings(Comp)
            Innings(Comp) = Innings(Comp + 1)
            Innings(Comp + 1) = tempname
            tempname = ERA(Comp)
            ERA(Comp) = ERA(Comp + 1)
            ERA(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Wins(J); Tab(32); Losses(J); Tab(39); Innings(J); Tab(49); Strikeouts(J); Tab(58); ERA(J)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Clemens.jpg")
picResults2.Cls
picResults2.Print "Roger Clemens lead the 2005 Houston"
picResults2.Print "Astros in strikeouts with 185."
End Sub

'This button uses a bubble sort to rearrange the players according to lowest earned run
'average.  It then reprints the data in this new sequence.  It also displays the picture
'of the player who lead the team in this statistical category as well as the earned run
'average he recorded in leading the team.
Private Sub cmdera_Click()
picResults.Cls
picResults.Print "Player"; Tab(25); "W"; Tab(32); "L"; Tab(39); "IP"; Tab(49); "K"; Tab(58); "ERA"
picResults.Print "----------"; Tab(25); "-----"; Tab(32); "-----"; Tab(39); "-----"; Tab(49); "-----"; Tab(58); "-----"
picResults.Print ""
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If ERA(Comp) > ERA(Comp + 1) Then
            tempname = ERA(Comp)
            ERA(Comp) = ERA(Comp + 1)
            ERA(Comp + 1) = tempname
            tempname = Player(Comp)
            Player(Comp) = Player(Comp + 1)
            Player(Comp + 1) = tempname
            tempname = Wins(Comp)
            Wins(Comp) = Wins(Comp + 1)
            Wins(Comp + 1) = tempname
            tempname = Innings(Comp)
            Innings(Comp) = Innings(Comp + 1)
            Innings(Comp + 1) = tempname
            tempname = Strikeouts(Comp)
            Strikeouts(Comp) = Strikeouts(Comp + 1)
            Strikeouts(Comp + 1) = tempname
            tempname = Losses(Comp)
            Losses(Comp) = Losses(Comp + 1)
            Losses(Comp + 1) = tempname
        End If
    Next Comp
Next Pass

For J = 1 To CTR
    picResults.Print Player(J); Tab(25); Wins(J); Tab(32); Losses(J); Tab(39); Innings(J); Tab(49); Strikeouts(J); Tab(58); ERA(J)
Next J
picPlayer.Picture = LoadPicture(App.Path & "\Clemens.jpg")
picResults2.Cls
picResults2.Print "Roger Clemens lead the 2005 Houston"
picResults2.Print "Astros in lowest ERA at 1.87."
End Sub

'This command button directs the user to the offensive statistics form.  It also clears
'the picture of the currently displayed statistical leader so that upon returning to the
'pitching form at a later time, all picture boxes will be clear.
Private Sub cmdoffense_Click()
    frmoffense.Show
    frmpitching.Hide
    picPlayer.Picture = LoadPicture("")
End Sub

'This command button directs the user to the main menu.  It also clears
'the picture of the currently displayed statistical leader so that upon returning to the
'pitching form at a later time, all picture boxes will be clear.
Private Sub cmdmain_Click()
    frmmain.Show
    frmpitching.Hide
    picPlayer.Picture = LoadPicture("")
End Sub

'This button allows the user to exit the program.
Private Sub cmdquit_Click()
End
End Sub


