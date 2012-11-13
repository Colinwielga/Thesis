VERSION 5.00
Begin VB.Form frmNuggets 
   BackColor       =   &H00FFFF00&
   Caption         =   "Denver Nuggets"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   Picture         =   "frmNuggets.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      Height          =   2295
      Left            =   3960
      ScaleHeight     =   2235
      ScaleWidth      =   5835
      TabIndex        =   4
      Top             =   960
      Width           =   5895
   End
   Begin VB.CommandButton cmdShowNuggets 
      BackColor       =   &H0000FFFF&
      Caption         =   "Show Denver Nuggets' Starters"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortPoints 
      BackColor       =   &H0000FFFF&
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdBackHome 
      BackColor       =   &H0000FFFF&
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label lblNuggets 
      BackColor       =   &H0000FFFF&
      Caption         =   "Denver Nuggets"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmNuggets"
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
frmNuggets.Hide

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShowNuggets_Click()
'open the data file
Open App.Path & "\NuggetsNotes.txt" For Input As #10

'get the data
ctr = 0
    Do While Not EOF(10)
        ctr = ctr + 1
        Input #10, player(ctr), points(ctr), rebounds(ctr), assists(ctr)
    Loop
    
picResults.Cls

'need to put in heading
picResults.Print "Starter"; Tab(27); "Points"; Tab(39); "Rebounds"; Tab(57); "Assists"
picResults.Print

'this will print the names and info
For I = 1 To ctr
    picResults.Print player(I); Tab(27); points(I); Tab(44); rebounds(I); Tab(59); assists(I)
Next I

cmdShowNuggets.Visible = False


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
'sort the players into alphabetical order
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
