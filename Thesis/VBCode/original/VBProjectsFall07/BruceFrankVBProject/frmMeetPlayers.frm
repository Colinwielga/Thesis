VERSION 5.00
Begin VB.Form frmMeetPlayers 
   BackColor       =   &H80000018&
   Caption         =   "Roster"
   ClientHeight    =   7125
   ClientLeft      =   2340
   ClientTop       =   2100
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10365
   Begin VB.CommandButton cmdBaller 
      Caption         =   "Click Here To See Player Info"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3480
      TabIndex        =   7
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   6120
      Width           =   3015
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5520
      ScaleHeight     =   2475
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton cmdPlayerInfo 
      Caption         =   "Click Me to Load Player Data"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtPlayerName 
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.PictureBox picRoster 
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   600
      ScaleHeight     =   4515
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdRoster 
      BackColor       =   &H00000000&
      Caption         =   "Click Here To See The Roster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblPlayerInfo 
      BackColor       =   &H80000009&
      Caption         =   "Which Player Would You Like to Know                            More About?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   6015
   End
   Begin VB.Image Image28 
      Height          =   1920
      Left            =   9000
      Picture         =   "frmMeetPlayers.frx":0000
      Top             =   5520
      Width           =   1530
   End
   Begin VB.Image Image27 
      Height          =   1845
      Left            =   7560
      Picture         =   "frmMeetPlayers.frx":0B1F
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Image Image26 
      Height          =   1800
      Left            =   6000
      Picture         =   "frmMeetPlayers.frx":16C8
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Image Image25 
      Height          =   1920
      Left            =   4560
      Picture         =   "frmMeetPlayers.frx":2D04
      Top             =   5400
      Width           =   1530
   End
   Begin VB.Image Image24 
      Height          =   1800
      Left            =   3000
      Picture         =   "frmMeetPlayers.frx":3823
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Image Image23 
      Height          =   1800
      Left            =   1440
      Picture         =   "frmMeetPlayers.frx":4E5F
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Image Image22 
      Height          =   1845
      Left            =   0
      Picture         =   "frmMeetPlayers.frx":649B
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Image Image21 
      Height          =   1920
      Left            =   9120
      Picture         =   "frmMeetPlayers.frx":7044
      Top             =   3600
      Width           =   1530
   End
   Begin VB.Image Image20 
      Height          =   1800
      Left            =   7440
      Picture         =   "frmMeetPlayers.frx":7B63
      Top             =   3600
      Width           =   1800
   End
   Begin VB.Image Image19 
      Height          =   1920
      Left            =   6000
      Picture         =   "frmMeetPlayers.frx":919F
      Top             =   3600
      Width           =   1530
   End
   Begin VB.Image Image18 
      Height          =   1800
      Left            =   4440
      Picture         =   "frmMeetPlayers.frx":9CBE
      Top             =   3600
      Width           =   1800
   End
   Begin VB.Image Image17 
      Height          =   1845
      Left            =   3000
      Picture         =   "frmMeetPlayers.frx":B2FA
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image Image16 
      Height          =   1920
      Left            =   1560
      Picture         =   "frmMeetPlayers.frx":BEA3
      Top             =   3600
      Width           =   1530
   End
   Begin VB.Image Image15 
      Height          =   1800
      Left            =   0
      Picture         =   "frmMeetPlayers.frx":C9C2
      Top             =   3600
      Width           =   1800
   End
   Begin VB.Image Image14 
      Height          =   1800
      Left            =   8880
      Picture         =   "frmMeetPlayers.frx":DFFE
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Image Image13 
      Height          =   1920
      Left            =   7440
      Picture         =   "frmMeetPlayers.frx":F63A
      Top             =   1800
      Width           =   1530
   End
   Begin VB.Image Image12 
      Height          =   1845
      Left            =   6000
      Picture         =   "frmMeetPlayers.frx":10159
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image Image11 
      Height          =   1800
      Left            =   4440
      Picture         =   "frmMeetPlayers.frx":10D02
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Image Image10 
      Height          =   1845
      Left            =   3120
      Picture         =   "frmMeetPlayers.frx":1233E
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image Image9 
      Height          =   1800
      Left            =   1440
      Picture         =   "frmMeetPlayers.frx":12EE7
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Image Image8 
      Height          =   1920
      Left            =   7800
      Picture         =   "frmMeetPlayers.frx":14523
      Top             =   0
      Width           =   1530
   End
   Begin VB.Image Image7 
      Height          =   1800
      Left            =   6120
      Picture         =   "frmMeetPlayers.frx":15042
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image6 
      Height          =   1845
      Left            =   4680
      Picture         =   "frmMeetPlayers.frx":1667E
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   1920
      Left            =   3240
      Picture         =   "frmMeetPlayers.frx":17227
      Top             =   0
      Width           =   1530
   End
   Begin VB.Image Image4 
      Height          =   1800
      Left            =   9240
      Picture         =   "frmMeetPlayers.frx":17D46
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   1920
      Left            =   0
      Picture         =   "frmMeetPlayers.frx":19382
      Top             =   1800
      Width           =   1530
   End
   Begin VB.Image Image2 
      Height          =   1845
      Left            =   1800
      Picture         =   "frmMeetPlayers.frx":19EA1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   0
      Picture         =   "frmMeetPlayers.frx":1AA4A
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmMeetPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Number(1 To 20) As String
Dim Name1(1 To 20) As String
Dim Position(1 To 20) As String
Dim Height1(1 To 20) As String
Dim Weight(1 To 20) As String
Dim School(1 To 20) As String
Dim Years(1 To 20) As String
Dim CTR As Integer



Private Sub cmdBaller_Click()
Dim PlayerName As String
PlayerName = txtPlayerName
'When i first looped through the array my resulting CTR tells me the number of values in my array so I can use my CTR as a restriction in the coming loops
'Found will be my variable which alerst me of a match
Dim Found As Boolean
'I need a new counter value since I can no longer use CTR
Dim Pos As Integer
Found = False
Do While Found = False And Pos < CTR
    Pos = Pos + 1
    'I will use an if function when searching for the matching name
    If PlayerName = Name1(Pos) Then
        Found = True
    End If
Loop
 
 'I must clear the picture box between each choice
 picPlayer.Cls
 'I will use another function to display the player info
 If Found = True Then
    picPlayer.Print "Number  "; Number(Pos)
    picPlayer.Print "Name  "; Name1(Pos)
    picPlayer.Print "Position  "; Position(Pos)
    picPlayer.Print "Height  "; Height1(Pos)
    picPlayer.Print "Weight  "; Weight(Pos)
    picPlayer.Print "School  "; School(Pos)
    picPlayer.Print "Years in the  NBA  "; Years(Pos)
 Else
    picPlayer.Print "Check your spelling!"
 End If
End Sub

Private Sub cmdPlayerInfo_Click()
'To sort the players into an array I must assign the various player info as variables.  The Variables included in Player info is Number, Name, Height, Position, Weight, School,and Years played in the NBA
'Now I Must sort the players into an array by using the outside file; Roster

Open App.Path & "\Roster.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Number(CTR), Name1(CTR), Position(CTR), Height1(CTR), Weight(CTR), School(CTR), Years(CTR)
Loop
Close #1

End Sub

Private Sub cmdreturn_Click()
'This allows the user to return to the Timberwolves main page
frmMeetPlayers.Visible = False
frmMainPage.Visible = True

End Sub

Private Sub cmdRoster_Click()

'This Shows the roster of the Minnesota Timberwolves
picRoster.Print "Corey Brewer"
picRoster.Print "Greg Buckner"
picRoster.Print "Michael Doleac"
picRoster.Print "Randy Foye"
picRoster.Print "Ryan Gomes"
picRoster.Print "Gerald Green"
picRoster.Print "Marko Jaric"
picRoster.Print "Al Jefferson"
picRoster.Print "Mark Madsen"
picRoster.Print "Rashad McCants"
picRoster.Print "Theo Ratliff"
picRoster.Print "Chris Richard"
picRoster.Print "Craig Smith"
picRoster.Print "Sebastian Telfair"
picRoster.Print "Antoine Walker"

End Sub

