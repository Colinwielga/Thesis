VERSION 5.00
Begin VB.Form frmPlayers 
   BackColor       =   &H80000012&
   Caption         =   "Players"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   Picture         =   "frmPlayers.frx":0000
   ScaleHeight     =   8685
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00008000&
      Height          =   6015
      Left            =   7200
      ScaleHeight     =   5955
      ScaleWidth      =   4635
      TabIndex        =   40
      Top             =   0
      Width           =   4695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmdPlayerRank 
      BackColor       =   &H00008000&
      Caption         =   "Find Player Ranks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7560
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00008000&
      Caption         =   "Clear Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdsort 
      BackColor       =   &H00008000&
      Caption         =   "Sort By Points Per Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00008000&
      Caption         =   "Find Total Points Per Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7200
      Width           =   1575
   End
   Begin VB.PictureBox picCasey 
      Height          =   2055
      Left            =   5760
      Picture         =   "frmPlayers.frx":452A
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   32
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picDupree 
      Height          =   2055
      Left            =   5760
      Picture         =   "frmPlayers.frx":5430
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   31
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox picBlount 
      Height          =   2055
      Left            =   5760
      Picture         =   "frmPlayers.frx":6304
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   30
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox PicCarter 
      Height          =   2055
      Left            =   0
      Picture         =   "frmPlayers.frx":7238
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   29
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picHassell 
      Height          =   2055
      Left            =   4320
      Picture         =   "frmPlayers.frx":81BC
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   28
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picReed 
      Height          =   2055
      Left            =   1440
      Picture         =   "frmPlayers.frx":904A
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   26
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picBanks 
      Height          =   2055
      Left            =   2880
      Picture         =   "frmPlayers.frx":9D1E
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   25
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picMarko 
      Height          =   2055
      Left            =   1440
      Picture         =   "frmPlayers.frx":AA18
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   24
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox picMadsen 
      Height          =   2055
      Left            =   4320
      Picture         =   "frmPlayers.frx":B7A7
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   23
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox picTroy 
      Height          =   2055
      Left            =   0
      Picture         =   "frmPlayers.frx":C60C
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   21
      Top             =   2400
      Width           =   1455
      Begin VB.PictureBox Picture8 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   855
         TabIndex        =   27
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   495
         Left            =   1440
         TabIndex        =   22
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox picBracey 
      Height          =   2055
      Left            =   2880
      Picture         =   "frmPlayers.frx":D52E
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   20
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox picRashad 
      Height          =   2055
      Left            =   4320
      Picture         =   "frmPlayers.frx":E2E5
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   19
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox picRicky 
      Height          =   2055
      Left            =   1440
      Picture         =   "frmPlayers.frx":F0F4
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox picKG 
      Height          =   2055
      Left            =   0
      Picture         =   "frmPlayers.frx":FE41
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox picEddie 
      Height          =   2055
      Left            =   2880
      Picture         =   "frmPlayers.frx":10C06
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00008000&
      Caption         =   "View All Players Numbers and Names and PPG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00008000&
      Caption         =   "Back To Home Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000006&
      Caption         =   "By: Chad Henfling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3720
      TabIndex        =   34
      Top             =   8160
      Width           =   2535
   End
   Begin VB.Label lblHeadCoach 
      BackColor       =   &H00FF0000&
      Caption         =   "Head Coach: Dwane Casey"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5880
      TabIndex        =   33
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label lblTroy 
      BackColor       =   &H00FF0000&
      Caption         =   "Troy Hudson"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblHassell 
      BackColor       =   &H00FF0000&
      Caption         =   "Trenton Hassell"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label lblDupree 
      BackColor       =   &H00FF0000&
      Caption         =   "Ronald Dupree"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblCarter 
      BackColor       =   &H00FF0000&
      Caption         =   "Anthony Carter"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label lblBlount 
      BackColor       =   &H00FF0000&
      Caption         =   "Mark Blount"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblBanks 
      BackColor       =   &H00FF0000&
      Caption         =   "Marcus Banks"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label lblMarko 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marko Jaric"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblMadsen 
      BackColor       =   &H00FF0000&
      Caption         =   "Mark Madsen"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblRashad 
      BackColor       =   &H00FF0000&
      Caption         =   "Rashad McCants"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblReed 
      BackColor       =   &H00FF0000&
      Caption         =   "Justin Reed"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label lblBracey 
      BackColor       =   &H00FF0000&
      Caption         =   "Bracey Wright"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblEddie 
      BackColor       =   &H00FF0000&
      Caption         =   "Eddie Griffin"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblRicky 
      BackColor       =   &H00FF0000&
      Caption         =   "Ricky Davis"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblKG 
      BackColor       =   &H00FF0000&
      Caption         =   "Kevin Garnett"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minnesota Timberwolves Center (MinnesotaTimberwovlesbyChadHenfling.vbp)
'Main Form (frmPlayers.frm)
'Chad Henfling
'Created March 23, 2006
'This form allows users to view all of the current Minnesota Timberwolves Players, Learn a bit about them, and sort them by player number, points per game, and find total team points per game.
Option Explicit
Private Sub cmdBack_Click()
    'goes back to main form
    frmPlayers.Visible = False
    frm1.Visible = True
End Sub

Private Sub cmdClear_Click()
    'clears picture box
    picOutput.Cls
End Sub


Private Sub cmdPlayerRank_Click()
    Dim score As Single
    Dim pass, pos, counter As Integer
    Dim temp As String
    Dim a, b As Single
    pos = 0
    a = 0
    pass = 0
    b = 0
    'first i will sort the players by points per game and then i will give them a rank.
    'bubble sort the points per game
    For pass = 1 To size
        For pos = 1 To (size - pass)
            If PPG(pos) > PPG(pos + 1) Then
                b = PPG(pos)
                PPG(pos) = PPG(pos + 1)
                PPG(pos + 1) = b
                temp = Names(pos)
                Names(pos) = Names(pos + 1)
                Names(pos + 1) = temp
                a = Numbers(pos)
                Numbers(pos) = Numbers(pos + 1)
                Numbers(pos + 1) = a
            End If
        Next pos
    Next pass
    For pos = 1 To size
    score = PPG(pos)
        'Shows the rank of players by points scored
        Select Case score
            Case Is > 15
                picOutput.Print "The Best Players Are        :  "; "ppg ="; PPG(pos); ","; Names(pos)
            Case 10 To 15
                picOutput.Print "The Good Players Are      :  "; "ppg= "; PPG(pos); ","; Names(pos)
            Case 5 To 10
                picOutput.Print "The Average Players Are :  "; "ppg="; PPG(pos); ","; Names(pos)
            Case Else
                picOutput.Print "Lowest Grade Players Are:  "; "ppg="; PPG(pos); ","; Names(pos)
        End Select
   Next pos
 
End Sub

Private Sub cmdQuit_Click()
    'When clicking on this rectangle gray button, you will currently and immedietly end the program from running its current processes.  This will not permanantly end the program, but close it until you wish to use it again.  It can be re-accessed by pressing the play button on the top of the screen.
    End
End Sub

Private Sub cmdsort_Click()
    Dim pass, pos As Integer
    Dim temp As String
    Dim a, b As Single
    pos = 0
    a = 0
    pass = 0
    b = 0
    
    'bubble sort the points per game
    For pass = 1 To size
        For pos = 1 To (size - pass)
            If PPG(pos) > PPG(pos + 1) Then
                b = PPG(pos)
                PPG(pos) = PPG(pos + 1)
                PPG(pos + 1) = b
                temp = Names(pos)
                Names(pos) = Names(pos + 1)
                Names(pos + 1) = temp
                a = Numbers(pos)
                Numbers(pos) = Numbers(pos + 1)
                Numbers(pos + 1) = a
            End If
        Next pos
    Next pass
    picOutput.Print "Sorted By Wolves Scoring leaders!"
    picOutput.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    For pos = 1 To size
        picOutput.Print PPG(pos), Names(pos); " # "; Numbers(pos)
    Next pos
End Sub

Private Sub cmdTotal_Click()
    Dim pos As Integer
    Dim total, temp As Single
    total = 0
    pos = 0
    'finds the teams total points per game by adding individual points per game.
    For pos = 1 To size
        temp = PPG(pos)
        total = total + temp
    Next pos
    'printing results in picture box
    picOutput.Print "*************************************************************************************************************************"
    picOutput.Print "The team's total points per game = "; total
End Sub

Private Sub cmdView_Click()
    Dim pass, pos As Integer
    Dim temp As String
    Dim b As Single
    Dim a As Integer
    'bubble sort of the scores
    For pass = 1 To size
        For pos = 1 To (size - pass)
            If Numbers(pos) > Numbers(pos + 1) Then
                a = Numbers(pos)
                Numbers(pos) = Numbers(pos + 1)
                Numbers(pos + 1) = a
                temp = Names(pos)
                Names(pos) = Names(pos + 1)
                Names(pos + 1) = temp
                b = PPG(pos)
                PPG(pos) = PPG(pos + 1)
                PPG(pos + 1) = b
            End If
        Next pos
    Next pass
    'printing the results
    For pos = 1 To size
        picOutput.Print Numbers(pos); Names(pos), , PPG(pos)
    Next pos
    
End Sub

'this part shows information players when you click on their picture.
Private Sub picBanks_Click()
    MsgBox "Marcus Banks, # 3 is a 6-2 gaurd from Nevada-Las Vegas.  He is now starting and averaging 8.8 points per gameand 3.4 assists per game.", , "Player Info"
End Sub

Private Sub picBlount_Click()
    mgsbox "Mark Blount, # 30 a 7-0 Center from University of Pittsburgh.  He is averaging 11.2 points per game.  He came to the Wolves half way through the season in a trade with the Boston Celtics.", , "Player Info"
End Sub

Private Sub picBracey_Click()
    MsgBox "Bracey Wright, # 6 a 6-3 guard from Indiana was averaging 3 points per game.  He is now playing in the NBA developmental league.", , "Player Info"
End Sub

Private Sub PicCarter_Click()
    MsgBox "Anthony Carter, # 17 is a 6-2 gaurd from Hawaii.  He is averaging 3.5 points per game.", , "Player Info"
End Sub

Private Sub picCasey_Click()
    MsgBox "Dwane Casey is the Head Coach of the Wolves.  This is his first season as a Head Coach.  This has been a rough first season for Casey and the Team with many ups and downs along with a losing record thus far.", , "Coach Info"
End Sub

Private Sub picDupree_Click()
    MsgBox "Ronald Dupree, # 12 from Louisiana State is averaging 2 points per game.  He is a 6-7 forward.", , "Player Info"
End Sub

Private Sub picEddie_Click()
    MsgBox "Eddie Griffin, # 41 a 6-10 Center is averaging 4.8 Points per game.  He is 240 pounds and went to Seton Hall.", , "Player Info"
End Sub

Private Sub picHassell_Click()
    MsgBox "Trenton Hassell, # 23 is a 6-5 Gaurd.  He is known for his outstanding defensive play.  As an added bonus he has scored 9.6 points per game this year for the Wolves.", , "Player Info"
End Sub

Private Sub picKG_Click()
    MsgBox "Kevin Garnett, # 21 also know as KG or the big ticket is the anchor for the wolves.  He is an 8 time NBA All-Star, 2003-04 NBA MVP, six-time All-NBA.  He was born in May 19th, 1976.  He is 6-11, 220 pounds.", , "Player Info"
End Sub

Private Sub picMadsen_Click()
    MsgBox "Mark 'MadDogg' Madsen, # 35 is a 6-9 center/forward from Stanford.  He is averaging 1.2 points per game, but is know for his hustle and determination he brings to the court.", , "Player Info"
End Sub

Private Sub picMarko_Click()
    MsgBox "Marko Jaric, # 55 is a 6-7 Guard from Serbia and Montenegro.  He is averaging 8.5 points per game.", , "Player Info"
End Sub

Private Sub picRashad_Click()
    MsgBox "Rashad McCants, # 1 is a 6-4 Guard.  He graduated from the national champion North Carolina Tar Heels last year.  He is averaging 6.8 points per game.", , "Player Info"
End Sub

Private Sub picReed_Click()
    MsgBox "Justin Reed, # 9 is a 6-8 forward from Mississippi.  He is averaging 3.6 points per game and 1.4 rebounds per game.", , "Player Info"
    
End Sub

Private Sub picRicky_Click()
    MsgBox "Ricky Davis, # 31 is a 6-7 Gaurd.  He is averaging 19.4 points per game.  He was born September 23rd, 1979 and graduated from Iowa University.", , "Player Info"
End Sub

Private Sub picTroy_Click()
    MsgBox "Troy Hudson, # 16 is a 6-1 Guard from Southern Illinois.  He weighs 170 pounds and is averaging 9.5 points per game.", , "Player Info"
End Sub

Private Sub Timer1_Timer()
    frmPlayers.Visible = False
    frm1.Visible = True
    Timer1 = False
End Sub
