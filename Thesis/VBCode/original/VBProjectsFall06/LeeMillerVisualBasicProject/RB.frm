VERSION 5.00
Begin VB.Form RB 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   Picture         =   "RB.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Tikicmd 
      BackColor       =   &H00000080&
      Caption         =   "Click to see Tiki Barber"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Statstxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   6720
      TabIndex        =   7
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton Playercmd 
      BackColor       =   &H00000080&
      Caption         =   "Find a few player's stats by entering a number from 1-3 in the text box"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton TDcmd 
      BackColor       =   &H00000080&
      Caption         =   "Sort Top 10 in Total Touchdowns"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Yardscmd 
      BackColor       =   &H00000080&
      Caption         =   "Sort Top 10 in Rushing Yards"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton RBcmd 
      BackColor       =   &H00000080&
      Caption         =   "Read Running Backs and then press another command"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Maincmd 
      BackColor       =   &H00000080&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label RBlabel 
      BackColor       =   &H00000080&
      Caption         =   "Running Backs"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "RB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name = Top 10 in 2005'
'Form Name = RB or Running Backs'
'Lee Miller
'Ocotober 31, 2006'
'The purpose of this form is to allow the user to search through
'a list of running backs, find the top 10 in yards and touchdowns.
'Then if the user wants to they can find the stats of 1, 2, or 3 players.
Dim RBName(1 To 100) As String
Dim RBTeam(1 To 100) As String
Dim RBYards(1 To 100) As Single
Dim RBTD(1 To 100) As Integer

Private Sub Tikicmd_Click()
'Loads a picture of Tiki Barber
picResults.Picture = LoadPicture(App.Path & "\Tiki1.bmp")
End Sub

Private Sub Form_Load()
'This makes the user have to press the first button before they can have acces to the
'others.
Yardscmd.Enabled = False
Playercmd.Enabled = False
TDcmd.Enabled = False
End Sub

Private Sub Maincmd_Click()
'Brings the user back to the main menu
picResults.Picture = LoadPicture("") 'Takes the picture of Tiki Barber away
Top10.Show
RB.Hide
Top10.Visible = True
End Sub

Private Sub Playercmd_Click()
'This button asks the player to type in a 1,2, or 3 and then type in names of the
'running backs they want to find and his stats will appear on the screen.
picResults.Picture = LoadPicture("") 'Takes the picture of Tiki Barber away
picResults.Cls
Dim Found As Boolean
Players = Val(Statstxt.Text)
If Players > 0 And Players < 4 Then
    picResults.Cls
    picResults.Print "Player"; Tab(30); "Team", "Yards", "Touchdowns"
    For I = 1 To Players
        A = InputBox("Enter the Name of a player", "Names")
        temp = 0
        Found = False
        Do While ((Not Found) And (temp < Counter))
            temp = temp + 1
            If A = RBName(temp) Then Found = True
        Loop
        If (Not Found) Then
            'Tells the user to try again'
            MsgBox "Player not found, try again", , "Error"
            I = I - 1
        Else
            'Prints the data of the names given
            picResults.Print RBName(temp); Tab(30); RBTeam(temp), RBYards(temp), RBTD(temp)
        End If
    Next I
Else
    'Tells the user to type in the right numbers'
    MsgBox "Please ente a number between 1 and 3 in the box.", , "Error"
End If
End Sub

Private Sub RBcmd_Click()
'This button reads the data which than allows the form to use it with the user'
Open App.Path & "\runningbacks.txt" For Input As #1
Counter = 0
Do While Not EOF(1)
    Counter = Counter + 1
    'puts the data into 4 arrays
    Input #1, RBName(Counter), RBTeam(Counter), RBYards(Counter), RBTD(Counter)
Loop
Close #1
'Allows the user to use these other buttons
Yardscmd.Enabled = True
Playercmd.Enabled = True
TDcmd.Enabled = True
End Sub

Private Sub TDcmd_Click()
'This button sorts through the players and gets the top 10 in Touchdowns'
picResults.Picture = LoadPicture("") 'Takes the picture of Tiki Barber away
For pass = 1 To Counter - 1
    For I = 1 To Counter - pass
        If RBTD(I) < RBTD(I + 1) Then
            Tempyards = RBYards(I)
            RBYards(I) = RBYards(I + 1)
            RBYards(I + 1) = Tempyards
            Tempteam = RBTeam(I)
            RBTeam(I) = RBTeam(I + 1)
            RBTeam(I + 1) = Tempteam
            Tempname = RBName(I)
            RBName(I) = RBName(I + 1)
            RBName(I + 1) = Tempname
            Temptd = RBTD(I)
            RBTD(I) = RBTD(I + 1)
            RBTD(I + 1) = Temptd
        End If
    Next I
Next pass
'Clears screen
picResults.Cls
'Prints a heading
picResults.Print "Player"; Tab(30); "Team", "Yards", "Touchdowns"
For I = 1 To 10
    'Prints the assortment
    picResults.Print RBName(I); Tab(30); RBTeam(I), RBYards(I), RBTD(I)
Next I
End Sub

Private Sub Yardscmd_Click()
'This button sorts through the players and gets the top 10 in Yards'
picResults.Picture = LoadPicture("") 'Takes the picture of Tiki Barber away
For pass = 1 To Counter - 1
    For I = 1 To Counter - pass
        If RBYards(I) < RBYards(I + 1) Then
            Tempyards = RBYards(I)
            RBYards(I) = RBYards(I + 1)
            RBYards(I + 1) = Tempyards
            Tempteam = RBTeam(I)
            RBTeam(I) = RBTeam(I + 1)
            RBTeam(I + 1) = Tempteam
            Tempname = RBName(I)
            RBName(I) = RBName(I + 1)
            RBName(I + 1) = Tempname
            Temptd = RBTD(I)
            RBTD(I) = RBTD(I + 1)
            RBTD(I + 1) = Temptd
        End If
    Next I
Next pass
'Clears screen
picResults.Cls
'Prints a heading
picResults.Print "Player"; Tab(30); "Team", "Yards", "Touchdowns"
For I = 1 To 10
    'Prints the assortment
    picResults.Print RBName(I); Tab(30); RBTeam(I), RBYards(I), RBTD(I)
Next I
End Sub
