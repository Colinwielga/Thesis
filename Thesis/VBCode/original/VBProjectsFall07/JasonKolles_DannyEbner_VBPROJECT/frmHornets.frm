VERSION 5.00
Begin VB.Form frmHornets 
   BackColor       =   &H00C00000&
   Caption         =   "New Orleans Hornets"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   Picture         =   "frmHornets.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Left            =   3960
      ScaleHeight     =   2235
      ScaleWidth      =   5835
      TabIndex        =   4
      Top             =   1080
      Width           =   5895
   End
   Begin VB.CommandButton cmdShowHornets 
      BackColor       =   &H00C0C000&
      Caption         =   "Show New Orleans Hornets' Starters"
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
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortPoints 
      BackColor       =   &H00C0C000&
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
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdBackHome 
      BackColor       =   &H00C0C000&
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
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C000&
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
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label lblHornets 
      BackColor       =   &H00C0C000&
      Caption         =   "New Orleans Hornets"
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
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmHornets"
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
frmHornets.Hide

End Sub

Private Sub cmdQuit_Click()
End
End Sub


Private Sub cmdShowHornets_Click()
'open the data file
Open App.Path & "\HornetsNotes.txt" For Input As #4

'get the data
ctr = 0
    Do While Not EOF(4)
        ctr = ctr + 1
        Input #4, player(ctr), points(ctr), rebounds(ctr), assists(ctr)
    Loop
    
picResults.Cls

'need to put in heading
picResults.Print "Starter"; Tab(27); "Points"; Tab(39); "Rebounds"; Tab(57); "Assists"
picResults.Print

'this will print the names and info
For I = 1 To ctr
    picResults.Print player(I); Tab(27); points(I); Tab(44); rebounds(I); Tab(59); assists(I)
Next I

cmdShowHornets.Visible = False

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
