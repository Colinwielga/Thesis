VERSION 5.00
Begin VB.Form frmhalloffame 
   BackColor       =   &H00400000&
   Caption         =   "Tecmo Hall of Fame"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdsortpos 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort By Position"
      Height          =   1935
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CmdSortteam 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort By Team"
      Height          =   1935
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdsortname 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort By Name"
      Height          =   1935
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Cmdsearch 
      BackColor       =   &H000000FF&
      Caption         =   "Search Player and See Profile"
      Height          =   1935
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdarray 
      BackColor       =   &H000000FF&
      Caption         =   "See TSB'S Best 15 Players"
      Height          =   1935
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox picresults 
      Height          =   5655
      Left            =   360
      ScaleHeight     =   5595
      ScaleWidth      =   9315
      TabIndex        =   1
      Top             =   2160
      Width           =   9375
   End
   Begin VB.CommandButton cmdmainmenu 
      BackColor       =   &H000000FF&
      Caption         =   "Main Menu"
      Height          =   1935
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmhalloffame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Super Tecmo Database
'Form name: frmhalloffame
'Author: Nate Johnson & Kevin Klein
'Date Written: October 11th, 2006
'Objective of project: This project will allow its users to learn more about the game of football
'and will also allow them the oppurtunity to learn how to play the game of football with the Nintendo
'video game, Tecmo Super Bowl.
'Objective of form: This form allows the user to see some of the top players in Tecmo Super Bowl,
'and allows the user to learn some background information on them. The players featured
'have been determined to be the best in the game, by the creators along with added opinions
'from prominent video game forum websites.

Option Explicit
Dim Player(1 To 15) As String
Dim Profile(1 To 15) As String
Dim Team(1 To 15) As String
Dim Spot(1 To 15) As String
Dim Counter As Integer
Dim I As Integer
Dim size As Integer

Private Sub cmdarray_Click() 'loads information into the array'
 picresults.Cls
    Counter = 0
    Open App.Path & "\hall of fame.txt" For Input As #2
    Do Until EOF(2)
        Counter = Counter + 1
        Input #2, Player(Counter), Team(Counter), Spot(Counter), Profile(Counter)
    Loop
    Close #2 'closes array'
    size = Counter
    picresults.Print "Player Name"; Tab(25); "Team"; Tab(60); "Position"
    picresults.Print "******************************************************************************************"
    For I = 1 To Counter
        picresults.Print Player(I); Tab(25); Team(I); Tab(50), Spot(I) 'prints loaded array data'
    Next I

End Sub


Private Sub cmdmainmenu_Click()
frmhalloffame.Hide 'hides hall of fame form'
frmMain.Show 'shows main form'

End Sub


Private Sub Cmdsearch_Click()
Dim search As String    'declares my variables
    Dim searchtrue As Boolean
    Dim pos As Integer
    Dim searchoutput As String
    searchtrue = False      'sets the variable for case where search isnt found'
    search = 0
    search = InputBox("Please enter the name of the character you are looking for.", "Character Name")      'opens a box so the user can input a search criteria'
    For pos = 1 To size     'loops through the list chekcing for the letters that the user searched for'
            searchoutput = InStr(Player(pos), search)
            If searchoutput <> 0 Then       'if something was found from the search it displays the match'
            searchtrue = True       'sets it so the search was successful
            MsgBox "The Player you selected is " & Player(pos) & ".  The profile: " & Profile(pos), , "Profile"     'displays the information to the user from their search
        End If
    Next pos
        If searchtrue = False Then      'loops through if the search was unsuccessful'
            MsgBox "No search results found, Please enter the name of a player on the displayed list'"
        End If
   
End Sub

Private Sub cmdsortname_Click() 'runs a sorting procedure
Dim pass, comp, J As Integer
Dim temp As String
Dim temp2 As String
Dim temp3 As String
Dim ctr As Single
ctr = 15

For pass = 1 To 14
    For comp = 1 To 15 - pass
        If Player(comp) > Player(comp + 1) Then
        temp = Player(comp)
        Player(comp) = Player(comp + 1)
        Player(comp + 1) = temp
        
                temp2 = Team(comp)
        Team(comp) = Team(comp + 1)
        Team(comp + 1) = temp2
        
                temp3 = Spot(comp)
        Spot(comp) = Spot(comp + 1)
        Spot(comp + 1) = temp3
        End If
    Next comp
Next pass

For J = 1 To 15
  picresults.Cls
 picresults.Print "Player Name"; Tab(25); "Team"; Tab(60); "Position"
    picresults.Print "******************************************************************************************"
    For I = 1 To Counter
        picresults.Print Player(I); Tab(25); Team(I); Tab(50), Spot(I) 'reprints the data in a new order
        Next I
        Next J
End Sub

Private Sub cmdsortpos_Click() 'runs a sorting procedure
Dim pass, comp, J As Integer
Dim temp As String
Dim temp2 As String
Dim temp3 As String
Dim ctr As Single
ctr = 15

For pass = 1 To 14
    For comp = 1 To 15 - pass
        If Spot(comp) > Spot(comp + 1) Then
        temp = Spot(comp)
        Spot(comp) = Spot(comp + 1)
        Spot(comp + 1) = temp
        
                temp2 = Player(comp)
        Player(comp) = Player(comp + 1)
        Player(comp + 1) = temp2
        
                temp3 = Player(comp)
        Player(comp) = Player(comp + 1)
        Player(comp + 1) = temp3
        End If
    Next comp
Next pass

For J = 1 To 15
  picresults.Cls
 picresults.Print "Player Name"; Tab(25); "Team"; Tab(60); "Position"
    picresults.Print "******************************************************************************************"
    For I = 1 To Counter
        picresults.Print Player(I); Tab(25); Team(I); Tab(50), Spot(I) 'reprints the data in a new order
        Next I
        Next J
End Sub

Private Sub CmdSortteam_Click() 'runs a sorting procedure
Dim pass, comp, J As Integer
Dim temp As String
Dim temp2 As String
Dim temp3 As String
Dim ctr As Single
ctr = 15

For pass = 1 To 14
    For comp = 1 To 15 - pass
        If Team(comp) > Team(comp + 1) Then
        temp = Team(comp)
        Team(comp) = Team(comp + 1)
        Team(comp + 1) = temp
        
                temp2 = Player(comp)
        Player(comp) = Player(comp + 1)
        Player(comp + 1) = temp2
        
                temp3 = Spot(comp)
        Spot(comp) = Spot(comp + 1)
        Spot(comp + 1) = temp3
        End If
    Next comp
Next pass

For J = 1 To 15
  picresults.Cls
 picresults.Print "Player Name"; Tab(25); "Team"; Tab(60); "Position"
    picresults.Print "******************************************************************************************"
    For I = 1 To Counter
        picresults.Print Player(I); Tab(25); Team(I); Tab(50), Spot(I) 'reprints the data in a new order
        Next I
        Next J
End Sub
