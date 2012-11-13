VERSION 5.00
Begin VB.Form frmWolves 
   BackColor       =   &H00FF0000&
   Caption         =   "Minnesota Timberwolves"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   Picture         =   "frmWolves.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdBackHome 
      BackColor       =   &H00008000&
      Caption         =   "Click Here to Go Back to Home Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSortPoints 
      BackColor       =   &H00008000&
      Caption         =   "Sort Starters by Average Points Scored Per Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdShowWolves 
      BackColor       =   &H00008000&
      Caption         =   "Show Minnesota Timberwolves Starters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      Height          =   2295
      Left            =   4080
      ScaleHeight     =   2235
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label lblWolves 
      BackColor       =   &H00FF0000&
      Caption         =   "Minnesota Timberwolves"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmWolves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim player(1 To 5) As String, points(1 To 5) As Integer
Dim rebounds(1 To 5) As Integer, assists(1 To 5) As Integer
Dim I As Integer, ctr As Integer
Dim ptsleader As String, mostpoints As Integer
Private Sub cmdBackHome_Click()
frmHome.Show
frmWolves.Hide

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShowWolves_Click()
'open the data file
Open App.Path & "\WolvesNotes.txt" For Input As #15

'get the data
ctr = 0
    Do While Not EOF(15)
        ctr = ctr + 1
        Input #15, player(ctr), points(ctr), rebounds(ctr), assists(ctr)
    Loop
    
picResults.Cls

'need to put in heading
picResults.Print "Starter"; Tab(27); "Points"; Tab(39); "Rebounds"; Tab(57); "Assists"
picResults.Print

'this will print the names and info
For I = 1 To ctr
    picResults.Print player(I); Tab(27); points(I); Tab(44); rebounds(I); Tab(59); assists(I)
Next I

cmdShowWolves.Visible = False



End Sub

Private Sub cmdSortPoints_Click()
'clear the box
picResults.Cls

'show the heading
picResults.Print "Starter"; Tab(27); "Points"; Tab(39); "Rebounds"; Tab(57); "Assists"
picResults.Print


'define some variables
Dim Pass As Integer, Pos As Integer
Dim Temp As String, temper As Integer
'sort the players by avg points per game
For Pass = 1 To ctr - 1
    For Pos = 1 To ctr - Pass
        If points(Pos) < points(Pos + 1) Then
            temper = points(Pos)
            points(Pos) = points(Pos + 1)
            points(Pos + 1) = temper
            
            Temp = player(Pos)
            player(Pos) = player(Pos + 1)
            player(Pos + 1) = Temp
                         
            temper = rebounds(Pos)
            rebounds(Pos) = rebounds(Pos + 1)
            rebounds(Pos + 1) = temper
            
            temper = assists(Pos)
            assists(Pos) = assists(Pos + 1)
            assists(Pos + 1) = temper
        End If
    Next Pos
Next Pass

'print sorted
For I = 1 To ctr
    picResults.Print player(I); Tab(27); points(I); Tab(44); rebounds(I); Tab(59); assists(I)
Next I
End Sub
