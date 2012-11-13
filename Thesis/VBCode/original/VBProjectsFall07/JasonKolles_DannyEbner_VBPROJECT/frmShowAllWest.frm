VERSION 5.00
Begin VB.Form frmShowAllWest 
   BackColor       =   &H000000FF&
   Caption         =   "All Western Conference Starters"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "frmShowAllWest.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   7335
      LargeChange     =   500
      Left            =   7560
      Max             =   8000
      SmallChange     =   100
      TabIndex        =   10
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdalpha 
      BackColor       =   &H00FF0000&
      Caption         =   "Show the Players in Alphabetical Order"
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
      Height          =   1575
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdload 
      BackColor       =   &H00FF0000&
      Caption         =   "Load the Players Data From a File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdastleader 
      BackColor       =   &H00FF0000&
      Caption         =   "Show the Assists Leader"
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
      Height          =   1575
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdrbleader 
      BackColor       =   &H00FF0000&
      Caption         =   "Show the Rebounds Leader"
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
      Height          =   1455
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdptsleader 
      BackColor       =   &H00FF0000&
      Caption         =   "Show the Points Leader"
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
      Height          =   1575
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdReadFiles 
      BackColor       =   &H00FF0000&
      Caption         =   "Show Western Conference Starters and Stats"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdBackHome 
      BackColor       =   &H00FF0000&
      Caption         =   "Click Here to Go to Home Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1695
   End
   Begin VB.PictureBox picResults2 
      Height          =   7280
      Left            =   3000
      ScaleHeight     =   7215
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      Begin VB.PictureBox picResults 
         AutoRedraw      =   -1  'True
         Height          =   15100
         Left            =   0
         ScaleHeight     =   15045
         ScaleWidth      =   4515
         TabIndex        =   11
         Top             =   0
         Width           =   4575
      End
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H000080FF&
      Caption         =   "Western Conference Starters' Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmShowAllWest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'input variables for program
Dim player(1 To 75) As String, points(1 To 75) As Integer
Dim rebounds(1 To 75) As Integer, assists(1 To 75) As Integer
Dim I As Integer, ctr As Integer
Dim ptsleader As String, mostpoints As Integer
Dim rbleader As String, mostreb As Integer
Dim astleader As String, mostast As Integer


Private Sub cmdalpha_Click()
'clear the box
picResults.Cls

'show the heading
picResults.Print "Starter"; Tab(27); "Points"; Tab(39); "Rebounds"; Tab(53); "Assists"
picResults.Print


'define some variables
Dim Pass As Integer, Pos As Integer
Dim Temp As String, temper As Integer
'sort the players into alphabetical order
For Pass = 1 To ctr - 1
    For Pos = 1 To ctr - Pass
        If player(Pos) > player(Pos + 1) Then
            Temp = player(Pos)
            player(Pos) = player(Pos + 1)
            player(Pos + 1) = Temp
            
            temper = points(Pos)
            points(Pos) = points(Pos + 1)
            points(Pos + 1) = temper
                         
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
    picResults.Print player(I); Tab(27); points(I); Tab(44); rebounds(I); Tab(55); assists(I)
Next I

cmdalpha.Visible = False


End Sub

Private Sub cmdastleader_Click()
'empty the picture box to show leaders
picResults.Cls
'Show assist leader
astleader = player(1)
mostast = assists(1)
For I = 2 To ctr
    If assists(I) > mostast Then
        mostast = assists(I)
        astleader = player(I)
    End If
Next I
picResults.Print astleader; " "; "lead the league"
picResults.Print "averaging"; mostast; "assists a game."

End Sub

Private Sub cmdBackHome_Click()
frmHome.Show
frmShowAllWest.Hide

End Sub

Private Sub cmdload_Click()
'open the data file
Open App.Path & "\AllWestNotes.txt" For Input As #21

'get the data
ctr = 0
    Do While Not EOF(21)
        ctr = ctr + 1
        Input #21, player(ctr), points(ctr), rebounds(ctr), assists(ctr)
    Loop
    
        'disable the load button after pushed
        
    cmdload.Visible = False
    cmdptsleader.Enabled = True
    cmdrbleader.Enabled = True
    cmdastleader.Enabled = True
    cmdalpha.Enabled = True
    cmdReadFiles.Enabled = True
    
    
    
    
    

End Sub

Private Sub cmdptsleader_Click()
'empty the picture box to show leaders
picResults.Cls
'show points leader
ptsleader = player(1)
mostpoints = points(1)
For I = 2 To ctr
    If points(I) > mostpoints Then
        mostpoints = points(I)
        ptsleader = player(I)
    End If
Next I
picResults.Print ptsleader; " "; "lead the league"
picResults.Print "averaging"; mostpoints; "points a game."
End Sub

Private Sub cmdQuit_Click()
End

End Sub

Private Sub cmdrbleader_Click()
'empty the picture box to show leaders
picResults.Cls
'show rebound leader
rbleader = player(1)
mostreb = rebounds(1)
For I = 2 To ctr
    If rebounds(I) > mostreb Then
        mostreb = rebounds(I)
        rbleader = player(I)
    End If
Next I
picResults.Print rbleader; " "; "lead the league"
picResults.Print "averaging"; mostreb; "rebounds a game."

End Sub

Private Sub cmdReadFiles_Click()

picResults.Cls

'need to put in heading
picResults.Print "Starter"; Tab(27); "Points"; Tab(39); "Rebounds"; Tab(53); "Assists"
picResults.Print

'this will print the names and info
For I = 1 To ctr
    picResults.Print player(I); Tab(27); points(I); Tab(44); rebounds(I); Tab(55); assists(I)
Next I

End Sub

Private Sub VScroll1_Change()
picResults.Top = -VScroll1.Value
End Sub
