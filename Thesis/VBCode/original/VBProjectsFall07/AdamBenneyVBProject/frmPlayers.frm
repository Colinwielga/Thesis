VERSION 5.00
Begin VB.Form frmPlayers 
   BackColor       =   &H000000FF&
   Caption         =   "Players Page"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack1 
      Caption         =   "Back to Main Screen"
      Height          =   975
      Index           =   1
      Left            =   11640
      TabIndex        =   12
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Screen"
      Height          =   735
      Index           =   0
      Left            =   17520
      TabIndex        =   11
      Top             =   11880
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoad2 
      Caption         =   "Load Full Roster"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11520
      TabIndex        =   10
      Top             =   5520
      Width           =   2055
   End
   Begin VB.PictureBox picTeam 
      Height          =   1575
      Left            =   9600
      ScaleHeight     =   1515
      ScaleWidth      =   5355
      TabIndex        =   9
      Top             =   3600
      Width           =   5415
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Starters"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   3975
   End
   Begin VB.CommandButton cmdSortNum 
      Caption         =   "Sort by Number"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   7
      Top             =   6120
      Width           =   4095
   End
   Begin VB.CommandButton cmdLookUpPlayer 
      Caption         =   "Look Up Player on Full Roster"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11520
      TabIndex        =   5
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortWeight 
      Caption         =   "Sort by Weight"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   7560
      Width           =   4095
   End
   Begin VB.CommandButton cmdSortName 
      Caption         =   "Sort by Name (Alphabetical)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   5400
      Width           =   4095
   End
   Begin VB.CommandButton cmdSortPos 
      Caption         =   "Sort by Position"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   6840
      Width           =   4095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Starters"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   3975
   End
   Begin VB.PictureBox picStarters 
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   1440
      Width           =   9135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "The Starters"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   735
      Left            =   3840
      TabIndex        =   6
      Top             =   480
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   10140
      Left            =   -1200
      Picture         =   "frmPlayers.frx":0000
      Top             =   0
      Width           =   20955
   End
End
Attribute VB_Name = "frmPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form: Players page (frmPlayers)
    'This form is an interactive form that shows the user the starting line-up for the St. John's lacrosse team and allows the user to sort the players by name, number, position, and weight.
    'The form also has the feature that allows the user to enter the name of a player he/she thinks is on the team but is not listed in the starting line-up.  If the player is on the roster, the player's name
    'number, and position will be shown.  An error message is shown if the name is not found on the roster.
    



Option Explicit
Dim jnumber(1 To 100) As Integer, pname(1 To 100) As String, ppos(1 To 100) As String, pyear(1 To 100) As String, pheight(1 To 100) As Single, pweight(1 To 100) As Single
Dim CTR As Integer, pos As Integer
Dim pass As Integer, tempname As String, tempnum As Integer, temppos As String, tempyear As String, tempheight As String, tempweight As Single
Dim CTR2 As Integer
Dim tplayer(1 To 100) As String, tpos(1 To 100) As String, tnumber(1 To 100) As Integer


Private Sub cmdBack1_Click(Index As Integer)        'This button allows the user to navigate back to the main (initial) screen.
frmFirst.Show
frmCoaches.Hide
frmPlayers.Hide
frmQuiz.Hide

End Sub

Private Sub cmdLoad_Click()                         'This button loads the starting line-up into an array.  Each player's name, jersey number, position, year, height and weight are loaded.
Open App.Path & "\starters.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, jnumber(CTR), pname(CTR), ppos(CTR), pyear(CTR), pheight(CTR), pweight(CTR)
Loop
Close #1
End Sub

Private Sub cmdLoad2_Click()                        'This button loads the entire roster in an array so that it can be searched.  Each player's name, position, and number are loaded.

Open App.Path & "\team.txt" For Input As #2
CTR2 = 0
Do While Not EOF(2)
    CTR2 = CTR2 + 1
    Input #2, tplayer(CTR2), tpos(CTR2), tnumber(CTR2)
Loop
Close #2
End Sub

Private Sub cmdLookUpPlayer_Click()                 'This button promts the user to enter the name of a player they believe is on the team but not listed as a starter.  The roster array is then searched for the name.
Dim found As Boolean, player As String
picTeam.Cls
player = InputBox("Enter name of play you think is on the team:", "PLAYER NAME")
pos = 0
found = False
Do While found = False And pos < CTR2
    pos = pos + 1
    If player = tplayer(pos) Then
        found = True
    End If
Loop
If found = False Then                               'If the player is not on the roster, an error message will appear.
    MsgBox player & " is not on the current roster.", , "NOT ON TEAM"
Else
    picTeam.Print "He is on the squad!"             'If the player is on the roster, his name, number and position will appear.
    picTeam.Print "Name: "; tplayer(pos), "Position: "; tpos(pos), "Number: "; tnumber(pos)
End If
End Sub

Private Sub cmdShow_Click()                         'This button makes the starting line-up, with each player's information, appear on the screen.
picStarters.Cls
picStarters.Print "No."; Tab(8); "Name"; Tab(30); "Pos."; Tab(50); "Year"; Tab(65); "Height"; Tab(80); "Weight"
picStarters.Print "*********************************************************************************************************"
For pos = 1 To CTR
    picStarters.Print jnumber(pos); Tab(8); pname(pos); Tab(30); ppos(pos); Tab(50); pyear(pos); Tab(65); pheight(pos); Tab(80); pweight(pos)
Next pos

End Sub

Private Sub cmdSortName_Click()                     'This button sorts the players alphabetically by last name.  It then swaps all the information for each play and arranges them to be shown in a picture box.
picStarters.Cls
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If pname(pos) > pname(pos + 1) Then
            tempname = pname(pos)
            pname(pos) = pname(pos + 1)
            pname(pos + 1) = tempname
            tempnum = jnumber(pos)
            jnumber(pos) = jnumber(pos + 1)
            jnumber(pos + 1) = tempnum
            temppos = ppos(pos)
            ppos(pos) = ppos(pos + 1)
            ppos(pos + 1) = temppos
            tempyear = pyear(pos)
            pyear(pos) = pyear(pos + 1)
            pyear(pos + 1) = tempyear
            tempheight = pheight(pos)
            pheight(pos) = pheight(pos + 1)
            pheight(pos + 1) = tempheight
            tempweight = pweight(pos)
            pweight(pos) = pweight(pos + 1)
            pweight(pos + 1) = tempweight
        End If
    Next pos
Next pass
picStarters.Print "No."; Tab(8); "Name"; Tab(30); "Pos."; Tab(50); "Year"; Tab(65); "Height"; Tab(80); "Weight"
picStarters.Print "*********************************************************************************************************"
For pos = 1 To CTR
    picStarters.Print jnumber(pos); Tab(8); pname(pos); Tab(30); ppos(pos); Tab(50); pyear(pos); Tab(65); pheight(pos); Tab(80); pweight(pos)
Next pos
End Sub

Private Sub cmdSortNum_Click()                  'This button sorts the players by number.  It then swaps all the information for each play and arranges them to be shown in a picture box.
picStarters.Cls
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If jnumber(pos) > jnumber(pos + 1) Then
            tempname = pname(pos)
            pname(pos) = pname(pos + 1)
            pname(pos + 1) = tempname
            tempnum = jnumber(pos)
            jnumber(pos) = jnumber(pos + 1)
            jnumber(pos + 1) = tempnum
            temppos = ppos(pos)
            ppos(pos) = ppos(pos + 1)
            ppos(pos + 1) = temppos
            tempyear = pyear(pos)
            pyear(pos) = pyear(pos + 1)
            pyear(pos + 1) = tempyear
            tempheight = pheight(pos)
            pheight(pos) = pheight(pos + 1)
            pheight(pos + 1) = tempheight
            tempweight = pweight(pos)
            pweight(pos) = pweight(pos + 1)
            pweight(pos + 1) = tempweight
        End If
    Next pos
Next pass
picStarters.Print "No."; Tab(8); "Name"; Tab(30); "Pos."; Tab(50); "Year"; Tab(65); "Height"; Tab(80); "Weight"
picStarters.Print "*********************************************************************************************************"
For pos = 1 To CTR
    picStarters.Print jnumber(pos); Tab(8); pname(pos); Tab(30); ppos(pos); Tab(50); pyear(pos); Tab(65); pheight(pos); Tab(80); pweight(pos)
Next pos
End Sub

Private Sub cmdSortPos_Click()                          'This button sorts the players by position.  It then swaps all the information for each play and arranges them to be shown in a picture box.
picStarters.Cls
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If ppos(pos) > ppos(pos + 1) Then
            tempname = pname(pos)
            pname(pos) = pname(pos + 1)
            pname(pos + 1) = tempname
            tempnum = jnumber(pos)
            jnumber(pos) = jnumber(pos + 1)
            jnumber(pos + 1) = tempnum
            temppos = ppos(pos)
            ppos(pos) = ppos(pos + 1)
            ppos(pos + 1) = temppos
            tempyear = pyear(pos)
            pyear(pos) = pyear(pos + 1)
            pyear(pos + 1) = tempyear
            tempheight = pheight(pos)
            pheight(pos) = pheight(pos + 1)
            pheight(pos + 1) = tempheight
            tempweight = pweight(pos)
            pweight(pos) = pweight(pos + 1)
            pweight(pos + 1) = tempweight
        End If
    Next pos
Next pass
picStarters.Print "No."; Tab(8); "Name"; Tab(30); "Pos."; Tab(50); "Year"; Tab(65); "Height"; Tab(80); "Weight"
picStarters.Print "*********************************************************************************************************"
For pos = 1 To CTR
    picStarters.Print jnumber(pos); Tab(8); pname(pos); Tab(30); ppos(pos); Tab(50); pyear(pos); Tab(65); pheight(pos); Tab(80); pweight(pos)
Next pos
End Sub

Private Sub cmdSortWeight_Click()                           'This button sorts the players by weight.  It then swaps all the information for each play and arranges them to be shown in a picture box.
Dim tempname5 As String, tempnum5 As Integer, temppos5 As String, tempyear5 As String, tempheight5 As Single, tempweight5 As Single
picStarters.Cls
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If pweight(pos) > pweight(pos + 1) Then
            tempname = pname(pos)
            pname(pos) = pname(pos + 1)
            pname(pos + 1) = tempname
            tempnum = jnumber(pos)
            jnumber(pos) = jnumber(pos + 1)
            jnumber(pos + 1) = tempnum
            temppos = ppos(pos)
            ppos(pos) = ppos(pos + 1)
            ppos(pos + 1) = temppos
            tempyear = pyear(pos)
            pyear(pos) = pyear(pos + 1)
            pyear(pos + 1) = tempyear
            tempheight = pheight(pos)
            pheight(pos) = pheight(pos + 1)
            pheight(pos + 1) = tempheight
            tempweight = pweight(pos)
            pweight(pos) = pweight(pos + 1)
            pweight(pos + 1) = tempweight
        End If
    Next pos
Next pass
picStarters.Print "No."; Tab(8); "Name"; Tab(30); "Pos."; Tab(50); "Year"; Tab(65); "Height"; Tab(80); "Weight"
picStarters.Print "*********************************************************************************************************"
For pos = 1 To CTR
    picStarters.Print jnumber(pos); Tab(8); pname(pos); Tab(30); ppos(pos); Tab(50); pyear(pos); Tab(65); pheight(pos); Tab(80); pweight(pos)
Next pos
End Sub
