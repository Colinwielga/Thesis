VERSION 5.00
Begin VB.Form QB 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   BeginProperty Font 
      Name            =   "Viner Hand ITC"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "QB.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox yardstxt 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   2640
      TabIndex        =   9
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Ycmd 
      BackColor       =   &H000000FF&
      Caption         =   "Find the players that passed for 4000, 3000, and 2000 yards"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox Statstxt 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   1095
      Left            =   6360
      TabIndex        =   7
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Playercmd 
      BackColor       =   &H000000FF&
      Caption         =   "Find a few player's stats by entering a number from 1-3 in the text box"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton TDcmd 
      BackColor       =   &H000000FF&
      Caption         =   "Sort Top 10 in Total touchdowns"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Yardscmd 
      BackColor       =   &H000000FF&
      Caption         =   "Sort Top 10 in Passing Yards"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Maincmd 
      BackColor       =   &H000000FF&
      Caption         =   "Main Menu"
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton QBcmd 
      BackColor       =   &H000000FF&
      Caption         =   "Read Quarterbacks and then press another command"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4275
      ScaleWidth      =   7395
      TabIndex        =   1
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Type in 4000, 3000, or 2000 in the box to the right and then hit the command button"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Quarterbacks"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "QB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name = Top 10 in 2005'
'Form Name = QB or Quarterbacks'
'Lee Miller
'Ocotober 31, 2006'
'The purpose of this form is to allow the user to search through
'a list of quarterbacks, find the top 10 in yards and touchdowns.
'Then if the user wants to they can find the stats of 1, 2, or 3 players.
Dim QBName(1 To 100) As String
Dim QBTeam(1 To 100) As String
Dim QBYards(1 To 100) As Single
Dim QBTD(1 To 100) As Integer

Private Sub Form_Load()
'This makes the user have to press the first button before they can have acces to the
'others.
Yardscmd.Enabled = False
Playercmd.Enabled = False
TDcmd.Enabled = False
Ycmd.Enabled = False
End Sub

Private Sub Maincmd_Click()
'Brings the user back to the main menu
Top10.Show
QB.Hide
Top10.Visible = True
End Sub

Private Sub Playercmd_Click()
'This button asks the user to type in a 1,2, or 3 and then type in names of the
'quarterbacks they want to find and his stats will appear on the screen.
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
            If A = QBName(temp) Then Found = True
        Loop
        If (Not Found) Then
            'Tells the user to try again'
            MsgBox "Player not found, try again", , "Error"
            I = I - 1
        Else
            'Prints the data of the names given
            picResults.Print QBName(temp); Tab(30); QBTeam(temp), QBYards(temp), QBTD(temp)
        End If
    Next I
Else
    'Tells the user to type in the right numbers'
    MsgBox "Please enter a number between 1 and 3 in the box.", , "Error"
End If
End Sub

Private Sub QBcmd_Click()
'This button reads the data which than allows the form to use it with the user'
Open App.Path & "\quarterbacks.txt" For Input As #1
Counter = 0
Do While Not EOF(1)
    Counter = Counter + 1
    'puts the data into 4 arrays
    Input #1, QBName(Counter), QBTeam(Counter), QBYards(Counter), QBTD(Counter)
Loop
Close #1
'Allows the user to use the other buttons
Yardscmd.Enabled = True
Playercmd.Enabled = True
TDcmd.Enabled = True
Ycmd.Enabled = True
End Sub

Private Sub TDcmd_Click()
'This button sorts through the players and gets the top 10 in Touchdowns'
For pass = 1 To Counter - 1
    For I = 1 To Counter - pass
        If QBTD(I) < QBTD(I + 1) Then
            Tempyards = QBYards(I)
            QBYards(I) = QBYards(I + 1)
            QBYards(I + 1) = Tempyards
            Tempteam = QBTeam(I)
            QBTeam(I) = QBTeam(I + 1)
            QBTeam(I + 1) = Tempteam
            Tempname = QBName(I)
            QBName(I) = QBName(I + 1)
            QBName(I + 1) = Tempname
            Temptd = QBTD(I)
            QBTD(I) = QBTD(I + 1)
            QBTD(I + 1) = Temptd
        End If
    Next I
Next pass
'Clears screen
picResults.Cls
'Prints a heading
picResults.Print "Player"; Tab(30); "Team", "Yards", "Touchdowns"
For I = 1 To 10
    'Prints the assortment
    picResults.Print QBName(I); Tab(30); QBTeam(I), QBYards(I), QBTD(I)
Next I
End Sub

Private Sub Yardscmd_Click()
'This button sorts through the players and gets the top 10 in Yards'
For pass = 1 To Counter - 1
    For I = 1 To Counter - pass
        If QBYards(I) < QBYards(I + 1) Then
            Tempyards = QBYards(I)
            QBYards(I) = QBYards(I + 1)
            QBYards(I + 1) = Tempyards
            Tempteam = QBTeam(I)
            QBTeam(I) = QBTeam(I + 1)
            QBTeam(I + 1) = Tempteam
            Tempname = QBName(I)
            QBName(I) = QBName(I + 1)
            QBName(I + 1) = Tempname
            Temptd = QBTD(I)
            QBTD(I) = QBTD(I + 1)
            QBTD(I + 1) = Temptd
        End If
    Next I
Next pass
'Clears screen
picResults.Cls
'Prints a heading
picResults.Print "Player"; Tab(30); "Team", "Yards", "Touchdowns"
For I = 1 To 10
    'Prints the assortment
    picResults.Print QBName(I); Tab(30); QBTeam(I), QBYards(I), QBTD(I)
Next I
End Sub

Private Sub Ycmd_Click()
'This feature allows the user to type in 2000, 3000, 4000 and then see all the
'passers that fall underneath each category.
picResults.Cls
Dim Y As Integer
Y = Val(yardstxt.Text) 'Gets value from the textbox
Select Case Y
    Case 4000 To 4999
        picResults.Print "4000 yard passers" 'Prints
        For I = 1 To 100
            If QBYards(I) >= Y Then
                picResults.Print QBName(I) 'Prints QB's that passed for 4000 yards or more
            End If
        Next I
     Case 3000 To 3999
        picResults.Print "3000 yard passers"
        For I = 1 To 100
            If QBYards(I) >= Y And QBYards(I) < (Y + 1000) Then
                picResults.Print QBName(I) 'Prints QB's that passed for 3000 yards or more
            End If
        Next I
     Case 2000 To 2999
        picResults.Print "2000 yard passers"
        For I = 1 To 100
            If QBYards(I) >= Y And QBYards(I) < (Y + 1000) Then
                picResults.Print QBName(I) 'Prints QB's that passed for 2000 yards or more
            End If
        Next I
     Case Else
        MsgBox "Type in 2000, 3000, or 4000 only", , "Invalid" 'Tells the user to do retry
End Select
End Sub
