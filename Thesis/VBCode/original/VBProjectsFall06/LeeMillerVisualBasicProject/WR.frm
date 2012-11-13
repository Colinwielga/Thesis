VERSION 5.00
Begin VB.Form WR 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   Picture         =   "WR.frx":0000
   ScaleHeight     =   8400
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Statstxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   6720
      TabIndex        =   7
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Playercmd 
      BackColor       =   &H00FFFF00&
      Caption         =   "Find a few player's stats by entering a number from 1-3 in the text box"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton TDcmd 
      BackColor       =   &H00FFFF00&
      Caption         =   "Sort Top 10 in Recieving Touchdowns"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Yardscmd 
      BackColor       =   &H00FFFF00&
      Caption         =   "Sort Top 10 in Recieving Yards"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton WRcmd 
      BackColor       =   &H00FFFF00&
      Caption         =   "Read Quarterbacks and then press another command"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Maincmd 
      BackColor       =   &H00FFFF00&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   1560
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Wide Recievers"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   48
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "WR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name = Top 10 in 2005'
'Form Name = WR or Wide Recievers'
'Lee Miller
'Ocotober 31, 2006'
'The purpose of this form is to allow the user to search through
'a list of wide recievers, find the top 10 in yards and touchdowns.
'Then if the user wants to they can find the stats of 1, 2, or 3 players.
Dim WRName(1 To 100) As String
Dim WRTeam(1 To 100) As String
Dim WRYards(1 To 100) As Single
Dim WRTD(1 To 100) As Integer

Private Sub Form_Load()
'This makes the user have to press the first button before they can have acces to the
'others.
Yardscmd.Enabled = False
Playercmd.Enabled = False
TDcmd.Enabled = False
End Sub

Private Sub Maincmd_Click()
'Brings the user back to the main menu
Top10.Show
QB.Hide
Top10.Visible = True
End Sub

Private Sub Playercmd_Click()
'This button asks the user to type in a 1,2, or 3 and then type in names of the
'wide recievers they want to find and his stats will appear on the screen.
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
            If A = WRName(temp) Then Found = True
        Loop
        If (Not Found) Then
            'Tells the user to try again'
            MsgBox "Player not found, try again", , "Error"
            I = I - 1
        Else
            'Prints the data of the names given
            picResults.Print WRName(temp); Tab(30); WRTeam(temp), WRYards(temp), WRTD(temp)
        End If
    Next I
Else
    'Tells the user to type in the right numbers'
    MsgBox "Please ente a number between 1 and 3 in the box.", , "Error"
End If
End Sub

Private Sub TDcmd_Click()
'This button sorts through the players and gets the top 10 in Touchdowns'
For pass = 1 To Counter - 1
    For I = 1 To Counter - pass
        If WRTD(I) < WRTD(I + 1) Then
            Tempyards = WRYards(I)
            WRYards(I) = WRYards(I + 1)
            WRYards(I + 1) = Tempyards
            Tempteam = WRTeam(I)
            WRTeam(I) = WRTeam(I + 1)
            WRTeam(I + 1) = Tempteam
            Tempname = WRName(I)
            WRName(I) = WRName(I + 1)
            WRName(I + 1) = Tempname
            Temptd = WRTD(I)
            WRTD(I) = WRTD(I + 1)
            WRTD(I + 1) = Temptd
        End If
    Next I
Next pass
'Clears screem
picResults.Cls
'Prints a heading
picResults.Print "Player"; Tab(30); "Team", "Yards", "Touchdowns"
For I = 1 To 10
    'Prints the assortment
    picResults.Print WRName(I); Tab(30); WRTeam(I), WRYards(I), WRTD(I)
Next I
End Sub

Private Sub WRcmd_Click()
'This button reads the data which than allows the form to use it with the user'
Open App.Path & "\widerecievers.txt" For Input As #1
Counter = 0
Do While Not EOF(1)
    Counter = Counter + 1
    'Puts the data into 4 arrays
    Input #1, WRName(Counter), WRTeam(Counter), WRYards(Counter), WRTD(Counter)
Loop
Close #1
'Allows the user to use the other buttons
Yardscmd.Enabled = True
Playercmd.Enabled = True
TDcmd.Enabled = True
End Sub

Private Sub Yardscmd_Click()
'This button sorts through the players and gets the top 10 in Yards'
For pass = 1 To Counter - 1
    For I = 1 To Counter - pass
        If WRYards(I) < WRYards(I + 1) Then
            Tempyards = WRYards(I)
            WRYards(I) = WRYards(I + 1)
            WRYards(I + 1) = Tempyards
            Tempteam = WRTeam(I)
            WRTeam(I) = WRTeam(I + 1)
            WRTeam(I + 1) = Tempteam
            Tempname = WRName(I)
            WRName(I) = WRName(I + 1)
            WRName(I + 1) = Tempname
            Temptd = WRTD(I)
            WRTD(I) = WRTD(I + 1)
            WRTD(I + 1) = Temptd
        End If
    Next I
Next pass
'Clears screen
picResults.Cls
'Prints a heading
picResults.Print "Player"; Tab(30); "Team", "Yards", "Touchdowns"
For I = 1 To 10
    'Prints the assortment
    picResults.Print WRName(I); Tab(30); WRTeam(I), WRYards(I), WRTD(I)
Next I
End Sub
