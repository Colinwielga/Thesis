VERSION 5.00
Begin VB.Form frmGrizzlies 
   BackColor       =   &H00800000&
   Caption         =   "Memphis Grizzlies"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   Picture         =   "frmGrizzlies.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Left            =   4200
      ScaleHeight     =   2235
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   1080
      Width           =   5775
   End
   Begin VB.CommandButton cmdShowGrizzlies 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Show Memphis Grizzlies' Starters"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortPoints 
      BackColor       =   &H00C0C0C0&
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdBackHome 
      BackColor       =   &H00C0C0C0&
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label lblGrizzlies 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Memphis Grizzlies"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmGrizzlies"
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
frmGrizzlies.Hide

End Sub

Private Sub cmdQuit_Click()
End
End Sub


Private Sub cmdShowGrizzlies_Click()
'open the data file
Open App.Path & "\GrizzliesNotes.txt" For Input As #3

'get the data
ctr = 0
    Do While Not EOF(3)
        ctr = ctr + 1
        Input #3, player(ctr), points(ctr), rebounds(ctr), assists(ctr)
    Loop
    
picResults.Cls

'need to put in heading
picResults.Print "Starter"; Tab(27); "Points"; Tab(39); "Rebounds"; Tab(57); "Assists"
picResults.Print

'this will print the names and info
For I = 1 To ctr
    picResults.Print player(I); Tab(27); points(I); Tab(44); rebounds(I); Tab(59); assists(I)
Next I

cmdShowGrizzlies.Visible = False


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