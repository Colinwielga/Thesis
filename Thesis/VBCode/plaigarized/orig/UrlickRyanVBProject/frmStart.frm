VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      TabIndex        =   8
      Top             =   3480
      Width           =   3135
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Handicap"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   7
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Scores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      TabIndex        =   6
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton cmdTerritory 
      Height          =   1575
      Left            =   120
      Picture         =   "frmStart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   3135
   End
   Begin VB.CommandButton cmdBlackberry 
      Height          =   735
      Left            =   120
      Picture         =   "frmStart.frx":2C4E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton cmdAlbany 
      Height          =   1095
      Left            =   120
      Picture         =   "frmStart.frx":51C7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblOr 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "-or-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCompute 
      Alignment       =   2  'Center
      BackColor       =   &H0001C5E7&
      Caption         =   "Compute your handicap"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      BackColor       =   &H0001C5E7&
      Caption         =   "Select a course to play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    

Private Sub cmdAlbany_Click()
    CourseName = "Albany"   'sets variable for dynamic picture loading
    
    Name1 = InputBox("Enter first player's name", "Player 1")   'input player(s) name(s)
    Name2 = InputBox("Enter second player's name (0 if no more players)", "Player 2")
    If Name2 = "0" Then
        frmStart.Hide
        frmFront.Show
    Else
        Name3 = InputBox("Please enter third player's name (0 if no more players)", "Player 3")
        If Name3 = "0" Then
            frmStart.Hide
            frmFront.Show
        Else
            Name4 = InputBox("Please enter fourth player's name (0 if no more players)", "Player 4")
        End If
    End If
    
    CTR = 0
    
    Open App.Path & "\Albany.txt" For Input As #1 'opens and loops through data file and stores course data into arrays
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Blue(CTR), White(CTR), Gold(CTR), Red(CTR), Par(CTR)
    Loop
    Close #1
    
    frmStart.Hide   'goes to next form
    frmFront.Show
End Sub

Private Sub cmdBlackberry_Click()
    CourseName = "Blackberry"   'sets variable for dynamic picture loading

    Name1 = InputBox("Enter first player's name", "Player 1")   'input player(s) name(s)
    Name2 = InputBox("Enter second player's name (0 if no player)", "Player 2")
    If Name2 = "0" Then
        frmStart.Hide
        frmFront.Show
    Else
        Name3 = InputBox("Please enter third player's name (0 if no player)", "Player 3")
        If Name3 = "0" Then
            frmStart.Hide
            frmFront.Show
        Else
            Name4 = InputBox("Please enter fourth player's name (0 if no player)", "Player 4")
        End If
    End If
    
    CTR = 0
    
    Open App.Path & "\Blackberry.txt" For Input As #1   'opens and loops through data file and stores course data into arrays
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Blue(CTR), White(CTR), Gold(CTR), Red(CTR), Par(CTR)
    Loop
    Close #1
    
    frmStart.Hide
    frmFront.Show
End Sub

Private Sub cmdCompute_Click()
Dim Pass As Integer
Dim Temp As Single

CTR = 0

Open App.Path & "\Differentials.txt" For Input As #1    'reads data from .txt file and stores into array
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, DiffArr(CTR)
Loop
Close #1

For Pass = 1 To CTR - 1                                 'sorts data in array from smallest to largest value
    For Pos = 1 To CTR - Pass
        If DiffArr(Pos) > DiffArr(Pos + 1) Then
            Temp = DiffArr(Pos)
            DiffArr(Pos) = DiffArr(Pos + 1)
            DiffArr(Pos + 1) = Temp
        End If
    Next Pos
Next Pass

HCap = DiffArr(1) * 0.96                                'takes lowest value from array and finsihes handicap computation

MsgBox "Your handicap is " & FormatNumber(HCap, 0), , "Your Handicap"   'displays computed handicap in message box

Erase DiffArr   'clears array for new data

CTR = 0
    
End Sub

Private Sub cmdEnter_Click()
    MsgBox "Information you will need to compute your handicap: 1. Your 5 most recent 18-hole scores 2. The course rating for each round 3. The course slope for each round"
    
    FirstScore = InputBox("Enter your most recent 18-hole score", "Enter First Score")              'user inputs data from a series of input boxes
    SecondScore = InputBox("Enter your second most recent 18-hole score", "Enter Second Score")
    ThirdScore = InputBox("Enter your third most recent 18-hole score", "Enter Third Score")
    FourthScore = InputBox("Enter your fourth most recent 18-hole score", "Enter Fourth Score")
    FifthScore = InputBox("Enter your fifth most recent 18-hole score", "Enter Fifth Score")
    
    frmStart.Hide       'switches forms
    frmHandicap.Show
    
End Sub

Private Sub cmdExit_Click()
    End     'exits program
End Sub

Private Sub cmdTerritory_Click()
    CourseName = "Territory"
    
    Name1 = InputBox("Enter first player's name", "Player 1")
    Name2 = InputBox("Enter second player's name (0 if no player)", "Player 2")
    If Name2 = "0" Then
        frmStart.Hide
        frmFront.Show
    Else
        Name3 = InputBox("Please enter third player's name (0 if no player)", "Player 3")
        If Name3 = "0" Then
            frmStart.Hide
            frmFront.Show
        Else
            Name4 = InputBox("Please enter fourth player's name (0 if no player)", "Player 4")
        End If
    End If
    
    CTR = 0
    
    Open App.Path & "\Territory.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Blue(CTR), White(CTR), Gold(CTR), Red(CTR), Par(CTR)
    Loop
    Close #1
    
    frmStart.Hide
    frmFront.Show
End Sub
