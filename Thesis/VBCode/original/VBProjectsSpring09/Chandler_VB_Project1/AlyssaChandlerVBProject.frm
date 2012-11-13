VERSION 5.00
Begin VB.Form frmHome1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "AlyssaChandlerVBProject.frx":0000
   ScaleHeight     =   8445
   ScaleMode       =   0  'User
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   14
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSamos 
      Caption         =   "Samos"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   12
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdRhodes 
      Caption         =   "Rhodes"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   11
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdHydra 
      Caption         =   "Hydra"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   10
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSantorini 
      Caption         =   "Santorini"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   9
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdCrete 
      Caption         =   "Crete"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdMykonos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mykonos"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   7080
      Width           =   1455
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2520
      ScaleHeight     =   2115
      ScaleWidth      =   4875
      TabIndex        =   6
      Top             =   4320
      Width           =   4935
   End
   Begin VB.CommandButton cmdPopulation 
      Caption         =   "Bustling or peaceful?"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdActive 
      Caption         =   "Active or Relaxing?"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdTemp 
      Caption         =   "Favorite Temperature"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdSeason 
      Caption         =   "Favorite Season"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblChooseIsland 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click your ideal island (listed above) for more information."
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   6600
      Width           =   6975
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Explore the Greek Isles"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Island Results"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose a search button to find the Greek island of your dreams! "
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label lblLine1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "What's Important to You on Vacation?"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   7215
   End
End
Attribute VB_Name = "frmHome1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim IslandName(1 To 100) As String, season(1 To 100) As String, AvgTemp(1 To 100) As Single, Active(1 To 100) As String, Population(1 To 100) As String, CTR3 As Integer

'Project Name: Ideal Greek Island
'Form: Home1
'Author: Alie Chandler
'Date Writen: started 3/11 finished 3/23
'Form Objective: This form is where the user will receive their ideal island/s based on their preferences. They will choose a button that is most important to them, and find islands that match their interests. Then they will choose an island button based on their results.
'Overall Project Objective: This project will match user's vacation interests with one or more Greek islands. The user can choose any island they want to learn information about the island, get hotel information and costs, and book their hotel.

Private Sub cmdActive_Click()
    Dim ActivityLevel As String, Found As Boolean, M As Integer
    CTR3 = 0
    'reads island array info file
    
    Open App.Path & "\Island.txt" For Input As #1
    
    Do Until EOF(1)
        CTR3 = CTR3 + 1
        Input #1, IslandName(CTR3), season(CTR3), AvgTemp(CTR3), Active(CTR3), Population(CTR3)
    Loop
    Close #1
    
    'exaustive search to match islands with users desire for activity on vacation
    picResults2.Cls
    
    Found = False
    M = 0
    ActivityLevel = InputBox("Do like an Active or Relaxing vacation?", "Activity Level")
    picResults2.Print "Your ideal island/s:"
    For M = 1 To CTR3
        If ActivityLevel = Active(M) Then
            Found = True
            picResults2.Print ; "=>"; IslandName(M)
        End If
    Next M
        
    picResults2.Print "...beacause of the "; ActivityLevel; " atmosphere."
    
    If (Not Found) Then
        picResults2.Cls
        MsgBox "Sorry, your answer is not an option. Choose Active or Relaxing.", , "Error"
    End If
    
End Sub

Private Sub cmdCrete_Click()
    'switch to different form by user's command
    frmCrete.Show
    frmHome1.Hide
End Sub

Private Sub cmdHydra_Click()
    'switch to different form by user's command
    frmHydra.Show
    frmHome1.Hide
End Sub

Private Sub cmdMykonos_Click()
    'switch to different form by user's command
    frmMykonos.Show
    frmHome1.Hide
End Sub

Private Sub cmdPopulation_Click()
    Dim PopLevel As String, Found As Boolean, M As Integer
    CTR3 = 0
    'reads island array info file
    
    Open App.Path & "\Island.txt" For Input As #1
    
    Do Until EOF(1)
        CTR3 = CTR3 + 1
        Input #1, IslandName(CTR3), season(CTR3), AvgTemp(CTR3), Active(CTR3), Population(CTR3)
    Loop
    Close #1
    picResults2.Cls
    
    'search array to match user's interests about amount of people around on vacations
    M = 0
    Found = False
    
    PopLevel = InputBox("Do you like interacting with a lot of people on vacation? (Yes or No)", "Population")
    picResults2.Print "Your ideal island/s:"
    
    For M = 1 To CTR3
        If PopLevel = Population(M) Then
            Found = True
            picResults2.Print "=>"; IslandName(M)
        End If
    Next M
    
    picResults2.Print "...because of the amount of people on the island."
    
    If (Not Found) Then
        picResults2.Cls
        MsgBox "Sorry, your answer is not an option. Please choose Yes or No.", , "Error"
    End If
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRhodes_Click()
    'switch to different form by user's command
    frmRhodes.Show
    frmHome1.Hide
End Sub

Private Sub cmdSamos_Click()
    'switch to different form by user's command
    frmSamos.Show
    frmHome1.Hide
End Sub

Private Sub cmdSantorini_Click()
    'switch to different form by user's command
    frmSantorini.Show
    frmHome1.Hide
End Sub

Private Sub cmdSeason_Click()
    Dim FavSeason As String, Found As Boolean, M As Integer
   
    CTR3 = 0
    'reads island array info file
    
    Open App.Path & "\Island.txt" For Input As #1
    
    Do Until EOF(1)
        CTR3 = CTR3 + 1
        Input #1, IslandName(CTR3), season(CTR3), AvgTemp(CTR3), Active(CTR3), Population(CTR3)
    Loop
    Close #1
    'exaustive search to match island with users favorite season
    picResults2.Cls
    
    M = 0
    Found = False
    
    FavSeason = InputBox("What is your favorite season to travel                          (Fall, Spring, Summer, or Winter)?", "Seasons")
    picResults2.Print "Your ideal island/s:"
    For M = 1 To CTR3
        If FavSeason = season(M) Then
            Found = True
            picResults2.Print "=>"; IslandName(M)
        End If
    Next M
    
    picResults2.Print "...because of its beautiful "; FavSeason
    
    If (Not Found) Then
        picResults2.Cls
        MsgBox "Season does not exist! Try again.", , "Error"
    End If
    
End Sub

Private Sub cmdTemp_Click()
    Dim FavTemp As Single
    'use static researched temperature data to match user's favorite temperature to an island
    picResults2.Cls
    
    FavTemp = InputBox("What's your ideal vacation temperature?", "Ideal Temperature")
    
    Select Case FavTemp
        Case Is >= 85
            picResults2.Print "Your ideal island is Rhodes during the summer."
        Case 75 To 84
            picResults2.Print "Your ideal island is Mykonos during the summer."
        Case 67 To 74
            picResults2.Print "Your ideal island is Crete during the spring."
        Case 61 To 66
            picResults2.Print "Your ideal island is Santorini during the fall."
        Case 55 To 60
            picResults2.Print "Your ideal island is Hydra during the spring."
        Case Is <= 54
            picResults2.Print "Your ideal island is Samos during the winter."
        Case Else
            MsgBox "Sorry, your answer is not an option. Enter an integer number.", , "Error"
    End Select
    
End Sub


