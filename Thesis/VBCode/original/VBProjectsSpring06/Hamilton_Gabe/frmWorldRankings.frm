VERSION 5.00
Begin VB.Form frmWorldRankings 
   BackColor       =   &H00C0FFC0&
   Caption         =   "World Rankings"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Britannic Bold"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmWorldRankings.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Player Stats"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0FFC0&
      FillColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   6480
      ScaleHeight     =   3075
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   1440
      Width           =   4455
   End
   Begin VB.PictureBox picRank 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6675
      ScaleWidth      =   5955
      TabIndex        =   8
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdEarnings05 
      BackColor       =   &H00C0FFC0&
      Caption         =   "2005 Earnings"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdEarningsC 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Career Earnings"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdWorldRank 
      BackColor       =   &H00C0FFC0&
      Caption         =   "WorldRank"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Gabe Hamilton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9600
      TabIndex        =   10
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label lblSort 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6960
      Width           =   1575
   End
End
Attribute VB_Name = "frmWorldRankings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WorldRank(1 To 100) As Integer
Dim FirstName(1 To 100), LastName(1 To 100) As String
Dim Earnings05(1 To 100), EarningsC(1 To 100) As Single
Dim size As Integer
Dim TempNum As Single
Dim TempName As String

Private Sub cmdEarnings05_Click()
    'Sorts information based on 2005 earnings
    Dim pass, pos As Integer
    For pass = 1 To (size - 1)
        For pos = 1 To (size - pass)
            If Earnings05(pos) < Earnings05(pos + 1) Then
            TempNum = Earnings05(pos)
            Earnings05(pos) = Earnings05(pos + 1)
            Earnings05(pos + 1) = TempNum
            TempName = FirstName(pos)
            FirstName(pos) = FirstName(pos + 1)
            FirstName(pos + 1) = TempName
            TempName = LastName(pos)
            LastName(pos) = LastName(pos + 1)
            LastName(pos + 1) = TempName
            TempNum = EarningsC(pos)
            EarningsC(pos) = EarningsC(pos + 1)
            EarningsC(pos + 1) = TempNum
            TempNum = WorldRank(pos)
            WorldRank(pos) = WorldRank(pos + 1)
            WorldRank(pos + 1) = TempNum
            End If
        Next pos
    Next pass
    picRank.Cls
    picRank.Print
    picRank.Print "World Rank"; Tab(15); "Name of Golfer"; Tab(40); "2005 Earnings"; Tab(60); "Career Earnings"
    picRank.Print "______________________________________________________________________________________________"
    picRank.Print
    For pos = 1 To size
        picRank.Print WorldRank(pos); Tab(15); FirstName(pos); Tab(25); LastName(pos); Tab(40); FormatCurrency(Earnings05(pos)); Tab(60); FormatCurrency(EarningsC(pos))
    Next pos
End Sub

Private Sub cmdEarningsC_Click()
    'Sorts information based on Career earnings
    Dim pass, pos As Integer
    For pass = 1 To (size - 1)
        For pos = 1 To (size - pass)
            If EarningsC(pos) < EarningsC(pos + 1) Then
            TempNum = EarningsC(pos)
            EarningsC(pos) = EarningsC(pos + 1)
            EarningsC(pos + 1) = TempNum
            TempName = FirstName(pos)
            FirstName(pos) = FirstName(pos + 1)
            FirstName(pos + 1) = TempName
            TempName = LastName(pos)
            LastName(pos) = LastName(pos + 1)
            LastName(pos + 1) = TempName
            TempNum = Earnings05(pos)
            Earnings05(pos) = Earnings05(pos + 1)
            Earnings05(pos + 1) = TempNum
            TempNum = WorldRank(pos)
            WorldRank(pos) = WorldRank(pos + 1)
            WorldRank(pos + 1) = TempNum
            End If
        Next pos
    Next pass
    picRank.Cls
    picRank.Print
    picRank.Print "World Rank"; Tab(15); "Name of Golfer"; Tab(40); "2005 Earnings"; Tab(60); "Career Earnings"
    picRank.Print "______________________________________________________________________________________________"
    picRank.Print
    For pos = 1 To size
        picRank.Print WorldRank(pos); Tab(15); FirstName(pos); Tab(25); LastName(pos); Tab(40); FormatCurrency(Earnings05(pos)); Tab(60); FormatCurrency(EarningsC(pos))
    Next pos
End Sub

Private Sub cmdHome_Click()
    'takes user back to the title page
    frmWorldRankings.Hide
    frmTitle.Show

End Sub

Private Sub cmdLoad_Click()
    'Loads information into picture box
    Dim pos As Integer
    pos = 0
    Open App.Path & "\WorldRank.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, WorldRank(pos), FirstName(pos), LastName(pos), Earnings05(pos), EarningsC(pos)
    Loop
    Close #1
    size = pos
    picRank.Cls
    picRank.Print
    picRank.Print "World Rank"; Tab(15); "Name of Golfer"; Tab(40); "2005 Earnings"; Tab(60); "Career Earnings"
    picRank.Print "______________________________________________________________________________________________"
    picRank.Print
    For pos = 1 To size
        picRank.Print WorldRank(pos); Tab(15); FirstName(pos); Tab(25); LastName(pos); Tab(40); FormatCurrency(Earnings05(pos)); Tab(60); FormatCurrency(EarningsC(pos))
    Next pos
    
End Sub

Private Sub cmdName_Click()
    'Sorts information based on golfer's name
    Dim pass, pos As Integer
    For pass = 1 To (size - 1)
        For pos = 1 To (size - pass)
            If FirstName(pos) > FirstName(pos + 1) Then
            TempName = FirstName(pos)
            FirstName(pos) = FirstName(pos + 1)
            FirstName(pos + 1) = TempName
            TempName = LastName(pos)
            LastName(pos) = LastName(pos + 1)
            LastName(pos + 1) = TempName
            TempNum = Earnings05(pos)
            Earnings05(pos) = Earnings05(pos + 1)
            Earnings05(pos + 1) = TempNum
            TempNum = EarningsC(pos)
            EarningsC(pos) = EarningsC(pos + 1)
            EarningsC(pos + 1) = TempNum
            TempNum = WorldRank(pos)
            WorldRank(pos) = WorldRank(pos + 1)
            WorldRank(pos + 1) = TempNum
            End If
        Next pos
    Next pass
    picRank.Cls
    picRank.Print
    picRank.Print "World Rank"; Tab(15); "Name of Golfer"; Tab(40); "2005 Earnings"; Tab(60); "Career Earnings"
    picRank.Print "______________________________________________________________________________________________"
    picRank.Print
    For pos = 1 To size
        picRank.Print WorldRank(pos); Tab(15); FirstName(pos); Tab(25); LastName(pos); Tab(40); FormatCurrency(Earnings05(pos)); Tab(60); FormatCurrency(EarningsC(pos))
    Next pos
End Sub

Private Sub cmdStats_Click()
    'user inputs golfer's name and statistical information appears
    Dim Residence, Education, LName, FName, Birthplace, Search As String
    Dim TurnedPro, JoinTour, NumberEvents, Birthdate, Birthyear, Rank As Integer
    Dim AvgDistance, DriveACC, GIR, AvgPutt, Earnings2005, EarningsCareer As Single
    Dim ScoringAvg, LongestDrive As Single
    Dim TopEvents, LowestScore As Integer
    Dim Found As Boolean
    Found = False
    Search = InputBox("Please Enter Name of Golfer", "Search PGA Golfers")
    Open App.Path & "\stats.txt" For Input As #1
    Do Until EOF(1) Or Found = True
        Input #1, FName, LName, Birthdate, Birthyear, Birthplace, Residence, Education, TurnedPro, JoinTour, NumberEvents, Rank, Earnings2005, EarningsCareer, AvgDistance, DriveACC, GIR, AvgPutt, ScoringAvg, TopEvents, LongestDrive, LowestScore
        Dim temp As String
        temp = FName + " " + LName
        If LCase(Search) = LCase(LName) Or LCase(Search) = LCase(FName) Or InStr(LCase(LName), LCase(Search)) <> 0 Or LCase(temp) = LCase(Search) Then
            Found = True
        End If
    Loop
    picResults.Cls
    If Found = True Then
        picResults.FontBold = True
        picResults.Print FName; " "; LName; Tab(20);
        picResults.FontBold = False
        picResults.Print "Rank: "; Rank; Tab(35); "Birthdate: "; Birthdate; "/"; Birthyear
        picResults.Print Tab(35); "Birthplace: "; Birthplace
        picResults.Print
        picResults.Print "Residence: "; Residence
        picResults.Print "Education: "; Education
        picResults.Print
        picResults.Print "Year Turned Pro: "; TurnedPro; Tab(30); "Events in 2005: "; NumberEvents
        picResults.Print "Year Joined Tour: "; JoinTour; Tab(30); "# of Top Ten Finishes: "; TopEvents
        picResults.Print
        picResults.Print "Avg Driving Dist: "; AvgDistance; Tab(35); "GIR: "; GIR; "%"
        picResults.Print "Driving Dist Acc: "; DriveACC; "%"; Tab(35); "Putting Avg: "; AvgPutt
        picResults.Print "Longest Drive: "; LongestDrive
        picResults.Print
        picResults.Print "Scoring Avg: "; ScoringAvg
        picResults.Print "Lowest Round: "; LowestScore
    Else
        picResults.Print "Sorry No Information"
    End If
    Close #1
End Sub

Private Sub cmdWorldRank_Click()
    'Sorts information based on World Rank
    Dim pass, pos As Integer
    For pass = 1 To (size - 1)
        For pos = 1 To (size - pass)
            If WorldRank(pos) > WorldRank(pos + 1) Then
            TempNum = WorldRank(pos)
            WorldRank(pos) = WorldRank(pos + 1)
            WorldRank(pos + 1) = TempNum
            TempName = FirstName(pos)
            FirstName(pos) = FirstName(pos + 1)
            FirstName(pos + 1) = TempName
            TempName = LastName(pos)
            LastName(pos) = LastName(pos + 1)
            LastName(pos + 1) = TempName
            TempNum = Earnings05(pos)
            Earnings05(pos) = Earnings05(pos + 1)
            Earnings05(pos + 1) = TempNum
            TempNum = EarningsC(pos)
            EarningsC(pos) = EarningsC(pos + 1)
            EarningsC(pos + 1) = TempNum
            End If
         Next pos
    Next pass
    picRank.Cls
    picRank.Print
    picRank.Print "World Rank"; Tab(15); "Name of Golfer"; Tab(40); "2005 Earnings"; Tab(60); "Career Earnings"
    picRank.Print "______________________________________________________________________________________________"
    picRank.Print
    For pos = 1 To size
        picRank.Print WorldRank(pos); Tab(15); FirstName(pos); Tab(25); LastName(pos); Tab(40); FormatCurrency(Earnings05(pos)); Tab(60); FormatCurrency(EarningsC(pos))
    Next pos
End Sub

