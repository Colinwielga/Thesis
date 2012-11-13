VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Roster"
   ClientHeight    =   7905
   ClientLeft      =   4200
   ClientTop       =   3450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10905
   Begin VB.PictureBox picSkierPhoto 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   7320
      ScaleHeight     =   2955
      ScaleWidth      =   3075
      TabIndex        =   7
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdSkierPhoto 
      BackColor       =   &H0080C0FF&
      Caption         =   "Get a photo of a skier"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   3495
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   3495
   End
   Begin VB.CommandButton cmdStartPage 
      BackColor       =   &H0080C0FF&
      Caption         =   "Back to Start Page"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   3495
   End
   Begin VB.CommandButton cmdPointSort 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sort by Points (Lowest = Best)"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton cmdAlphabetSort 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sort Alphabetically"
      Height          =   735
      Left            =   240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   3495
   End
   Begin VB.CommandButton cmdYearSort 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sort by year in school"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   3495
   End
   Begin VB.PictureBox picDisplayNames 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   4320
      ScaleHeight     =   3675
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Label lblYears 
      BackColor       =   &H00FFFFC0&
      Caption         =   "2008-2009"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblRoster 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Roster"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   4320
      TabIndex        =   8
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   240
      Picture         =   "frmRoster.frx":0000
      Top             =   240
      Width           =   3660
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project: SJU_Ski_Team
'Form Name: frmRoster
'Author: Kevin Neal
'Written: March 21, 2009
'Object: 1)Read a file with names, grades, and CCSA points(needing module/public variables)
        '2)Sort alphabetically, by grade, or by CCSA rank with parallel arrays
        '3)Load pictures using and input from and input box, boolean variables,
        '  and an array search

 'Declare variables use for sorting
    Dim Pass As Integer, Pos As Integer, TempName As String
    Dim TempGrade As String, TempNum As Integer, TempScore As Single

Private Sub cmdAlphabetSort_Click()

    'Sort Skiers by name
    For Pass = 1 To (SkierCTR - 1)                          'Keep track of passes
        For Pos = 1 To (SkierCTR - Pass)                    'Keep track of comparisons
            If SkierNames(Pos) > SkierNames(Pos + 1) Then   'Exchange names
                TempName = SkierNames(Pos)
                SkierNames(Pos) = SkierNames(Pos + 1)
                SkierNames(Pos + 1) = TempName
                TempGrade = SkierGrades(Pos)                'Exchange grades
                SkierGrades(Pos) = SkierGrades(Pos + 1)
                SkierGrades(Pos + 1) = TempGrade
                TempNum = NumGrade(Pos)
                NumGrade(Pos) = NumGrade(Pos + 1)
                NumGrade(Pos + 1) = TempNum
                TempScore = SkierScore(Pos)                 'Exchange Scores
                SkierScore(Pos) = SkierScore(Pos + 1)
                SkierScore(Pos + 1) = TempScore
            End If
        Next Pos
    Next Pass
    
    'Clear Screen and get a header
    picDisplayNames.Cls
    picDisplayNames.Print "Name"; Tab(30); "Grade"; Tab(50); "Points"
    picDisplayNames.Print "============================================================"
   
    
    
    'Print the Results
    'Make a counter variable
    Dim I As Integer
    For I = 1 To SkierCTR Step 1
        picDisplayNames.Print SkierNames(I); Tab(30); SkierGrades(I); Tab(50); FormatNumber(SkierScore(I), 2)
    Next I
        
End Sub


Private Sub cmdPointSort_Click()
    'Sort By points
    
    For Pass = 1 To (SkierCTR - 1)                          'Keep track of passes
        For Pos = 1 To (SkierCTR - Pass)                    'Keep track of comparisons
            If SkierScore(Pos) > SkierScore(Pos + 1) Then
                TempName = SkierNames(Pos)                  'Exchange names
                SkierNames(Pos) = SkierNames(Pos + 1)
                SkierNames(Pos + 1) = TempName
                TempGrade = SkierGrades(Pos)                'Exchange grades
                SkierGrades(Pos) = SkierGrades(Pos + 1)
                SkierGrades(Pos + 1) = TempGrade
                TempNum = NumGrade(Pos)
                NumGrade(Pos) = NumGrade(Pos + 1)
                NumGrade(Pos + 1) = TempNum
                TempScore = SkierScore(Pos)                 'Exchange Scores
                SkierScore(Pos) = SkierScore(Pos + 1)
                SkierScore(Pos + 1) = TempScore
            End If
        Next Pos
    Next Pass
    
     'Clear Screen and get a header
    picDisplayNames.Cls
    picDisplayNames.Print "Name"; Tab(30); "Grade"; Tab(50); "Points"
    picDisplayNames.Print "============================================================"
    
    'Print the Results
    'Make a counter variable
    Dim K As Integer
    For K = 1 To SkierCTR Step 1
        picDisplayNames.Print SkierNames(K); Tab(30); SkierGrades(K); Tab(50); FormatNumber(SkierScore(K), 2)
    Next K
End Sub

Private Sub cmdQuit_Click()
    'Quit program
    End
End Sub

Private Sub cmdSkierPhoto_Click()
    'Display the photo of the skier with input from input box
    'Load pictures with a picture box
    
    'Declare and initialize variables
    Dim SkierID As String, Found As Boolean, I As Integer, LastName As String
    I = 1
    SkierID = InputBox("Who would you like to know more about? (Last name, First name)", "Input Name")
    
    Do While (Not Found And I < SkierCTR)
        If SkierID = SkierNames(I) Then
            Found = True
        Else
            I = I + 1
        End If
    Loop
    
    If Found = True Then
        picSkierPhoto.Picture = LoadPicture(App.Path & "\" & SkierID & ".jpeg")
    Else
        MsgBox "Sorry, no one on the team is named " & SkierID & ".  You may want to check your spelling"
    End If
    
        
        
End Sub


Private Sub cmdStartPage_Click()
    'Go back to the home screen
    frmRoster.Hide 'Hide this form
    frmStartPage.Show 'Show Start Page
    
    'Clear picture
    picSkierPhoto.Picture = LoadPicture(App.Path & "\Blank.jpg")
    
End Sub

Private Sub cmdYearSort_Click()
    
    'Sort Skiers by Grade
    For Pass = 1 To (SkierCTR - 1)                          'Keep track of passes
        For Pos = 1 To (SkierCTR - Pass)                    'Keep track of comparisons
            If NumGrade(Pos) > NumGrade(Pos + 1) Then
                TempName = SkierNames(Pos)                  'Exchange names
                SkierNames(Pos) = SkierNames(Pos + 1)
                SkierNames(Pos + 1) = TempName
                TempGrade = SkierGrades(Pos)                'Exchange grades
                SkierGrades(Pos) = SkierGrades(Pos + 1)
                SkierGrades(Pos + 1) = TempGrade
                TempNum = NumGrade(Pos)
                NumGrade(Pos) = NumGrade(Pos + 1)
                NumGrade(Pos + 1) = TempNum
                TempScore = SkierScore(Pos)                 'Exchange Scores
                SkierScore(Pos) = SkierScore(Pos + 1)
                SkierScore(Pos + 1) = TempScore
            End If
        Next Pos
    Next Pass
    
    'Clear Screen and get a header
    picDisplayNames.Cls
    picDisplayNames.Print "Name"; Tab(30); "Grade"; Tab(50); "Points"
    picDisplayNames.Print "============================================================"
    
    'Print the Results
    'Make a counter variable
    Dim J As Integer
    For J = 1 To SkierCTR Step 1
        picDisplayNames.Print SkierNames(J); Tab(30); SkierGrades(J); Tab(50); FormatNumber(SkierScore(J), 2)
    Next J
End Sub

