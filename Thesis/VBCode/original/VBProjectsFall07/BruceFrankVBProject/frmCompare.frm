VERSION 5.00
Begin VB.Form frmCompare 
   Caption         =   "Compare the T-Wolves"
   ClientHeight    =   8715
   ClientLeft      =   2550
   ClientTop       =   1275
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Myriad Condensed Web"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   10455
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Retun to Main Page"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   18
      Top             =   6120
      Width           =   1695
   End
   Begin VB.PictureBox picAscending 
      BackColor       =   &H80000009&
      Height          =   8295
      Left            =   8520
      ScaleHeight     =   8235
      ScaleWidth      =   1755
      TabIndex        =   17
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Click here to see NBA rankings from worst to best (That way the Timberwolves are a little closer to the top)"
      Height          =   1215
      Left            =   5760
      TabIndex        =   16
      Top             =   2400
      Width           =   2655
   End
   Begin VB.PictureBox picCompare 
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   2955
      TabIndex        =   15
      Top             =   7200
      Width           =   3015
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Enter the team you want and press here to compare"
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox txtcompareteam 
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H000000FF&
      Caption         =   "Click this to load the data"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblchoice10 
      BackColor       =   &H00C00000&
      Caption         =   "Sacramento"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label lblchoice9 
      BackColor       =   &H00C00000&
      Caption         =   "Milwaukee    Seattle    Memphis"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label lblchoice8 
      BackColor       =   &H00C00000&
      Caption         =   "Philadelphia   Charlotte   New York"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Label lblchoice7 
      BackColor       =   &H00C00000&
      Caption         =   "Denver   New Jersey  New Orleans"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label lblchoice6 
      BackColor       =   &H00C00000&
      Caption         =   "Detroit   Chicago   Golden State"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label lblchoice5 
      BackColor       =   &H00C00000&
      Caption         =   "Houston  Phoenix   L.A. Clippers"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label lblChoice4 
      BackColor       =   &H00C00000&
      Caption         =   "Indiana   San Antonio   Boston  "
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label lblChoice3 
      BackColor       =   &H00C00000&
      Caption         =   "Toronto    Atlanta     Portland"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label lblchoice2 
      BackColor       =   &H00C00000&
      Caption         =   "Dallas    Utah   Miami    Orlando"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label lblChoices 
      BackColor       =   &H00C00000&
      Caption         =   "L.A. Lakers   Cleveland    Washington"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lblList 
      BackColor       =   &H00C00000&
      Caption         =   "Team Choices:"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblCompare 
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "How do the T-Wolves stack up against other NBA teams? "
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -960
      Picture         =   "frmCompare.frx":0000
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'During this form I will use a file input and arrays
'i dim the following variables as arrays and under option explicit because i need them assecible on the form level
Dim TeamName(1 To 31) As String
Dim TeamRanking(1 To 31) As Single
Dim CTR As Integer


Private Sub cmdCompare_Click()
'This command button processes array information with a calculation and displays the output
'the user is asked to enter a team name in a text box which they want to compare with the timberwolves.  The program will then inform the user how much higher or lower ranked the team is compared to the timberwolves
'This is essentially a two part program, 1.match and stop and 2.processing arrays with a calculation
'Match and stop finds the team the user wants while the calculation compares the teams rank to the timberwolves rank

Dim Space As Single
Dim Team As String
Dim Pos As Integer
Dim Found As Boolean
Found = False
Team = txtcompareteam
picCompare.Cls
Pos = 0
Do While Found = False And Pos < CTR
    Pos = Pos + 1
    If Team = TeamName(Pos) Then
        Found = True
    End If
Loop
If TeamRanking(Pos) < 28 Then
    Space = 28 - TeamRanking(Pos)
    picCompare.Print TeamName(Pos); " is ranked "; Space
    picCompare.Print "higher than the Timberwolves"
End If
If TeamRanking(Pos) > 28 Then
    Space = TeamRanking(Pos) - 28
    picCompare.Print TeamName(Pos); " is ranked "; Space
    picCompare.Print " lower than the Timberwolves"
End If

End Sub

Private Sub cmdLoad_Click()
'this button loads the file into an array.  Its useful to load the data with one button and then use other buttons to manipulate the data because you don't have to continnually reload data.
'Once the text file is loaded once it can be accesed time and again by outside command buttons

Open App.Path & "\nbastandings.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, TeamName(CTR), TeamRanking(CTR)
Loop
Close #1

End Sub

Private Sub cmdreturn_Click()
'This button returns the user to the main page form by hiding the compare form.
frmCompare.Visible = False
frmMainPage.Visible = True

End Sub

Private Sub cmdReverse_Click()
'This button sorts the array by rank.  It will show the worst NBA team to the best
'Although it sorts from worse to best, the rank numbers themselves are actually being sorted in descending order because 30 is the worst team down to 1 which is the best

Dim Pos As Integer
Dim Temp As Single
Dim Temp1 As String
Dim Pass As Integer
Dim Comp As Integer
For Pass = 1 To (CTR - 1)
    For Comp = 1 To (CTR - Pass)
        'this portion sorts the variables into their new positions
        If TeamRanking(Comp) < TeamRanking(Comp + 1) Then
            Temp = TeamRanking(Comp)
            TeamRanking(Comp) = TeamRanking(Comp + 1)
            TeamRanking(Comp + 1) = Temp
            
            Temp1 = TeamName(Comp)
            TeamName(Comp) = TeamName(Comp + 1)
            TeamName(Comp + 1) = Temp1
        End If
    Next Comp
Next Pass
'This prints the output
For Pos = 1 To CTR
    picAscending.Print TeamName(Pos)
Next Pos




        
        
End Sub
