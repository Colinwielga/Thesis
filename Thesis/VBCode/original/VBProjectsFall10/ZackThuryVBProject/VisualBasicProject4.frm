VERSION 5.00
Begin VB.Form FrmTopGolfers 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   11685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13440
   FillColor       =   &H00FF0000&
   ForeColor       =   &H80000001&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11685
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9720
      Width           =   2175
   End
   Begin VB.CommandButton cmdBackToHome 
      BackColor       =   &H00808000&
      Caption         =   "Back To Home Screen"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9720
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortByCountry 
      Caption         =   "Sort Golfers By Country"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   5
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortByYearTurnedPro 
      BackColor       =   &H00808080&
      Caption         =   "Sort Golfers By Year Turned Pro"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      TabIndex        =   4
      Top             =   9720
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortByWins 
      Caption         =   "Sort Golfers By Wins"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      TabIndex        =   3
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortByName 
      Caption         =   "Sort Golfers By Name"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   2
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton cmdShowTop10 
      Caption         =   "Show Top 15 Golfers"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   8760
      Width           =   2175
   End
   Begin VB.PictureBox picTop15 
      BackColor       =   &H00FFFFFF&
      Height          =   6255
      Left            =   1680
      ScaleHeight     =   6195
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   1920
      Width           =   9855
   End
   Begin VB.Label lblTopRanked 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Top 15 Ranked Professional Golfers on the PGA Tour:"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "FrmTopGolfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'prints the file into a picture box and then sorts the content by different things such as the golfer name, wins, etc.
'declares variables
Dim GolferName(1 To 15) As String
Dim Country(1 To 15) As String
Dim PGATourWins(1 To 15) As Integer
Dim YearTurnedPro(1 To 15) As Integer
Dim Rank(1 To 15) As Integer
Dim Counter As Single
Dim Pass As Integer
Dim Pos As Integer
Dim TempGolferName As String
Dim TempCountry As String
Dim TempPGATourWins As Integer
Dim TempYearTurnedPro As Integer
Dim TempRank As Integer
Dim I As Integer

'hides the topgolfers form and lets the user see the title form
Private Sub cmdBackToHome_Click()
    FrmTopGolfers.Hide
    FrmTitle.Show
End Sub

'quits the program
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdShowTop10_Click()
    'prints the headings and titles for the table of the top 15 professional golfers
    picTop15.Cls
    picTop15.Print "Rank"; Tab(10); "Golfer Name"; Tab(40); "Country"; Tab(65); "Year Turned Pro"; Tab(95); "PGA Tour Wins"
    picTop15.Print "_________________________________________________________________________________________________________"
    'opens top15golfers file from same folder that the project is saved in and saves the variable to arrays
    Open App.Path & "\Top15Golfers.txt" For Input As #1
    Counter = 0
    'prints the contents of the file into the picture box
    Do While Not EOF(1)
        Counter = Counter + 1
        Input #1, Rank(Counter), GolferName(Counter), Country(Counter), YearTurnedPro(Counter), PGATourWins(Counter)
        picTop15.Print Tab(2); Rank(Counter); Tab(10); GolferName(Counter); Tab(40); Country(Counter); Tab(69); YearTurnedPro(Counter); Tab(100); PGATourWins(Counter)
    Loop
    'closes the file
    Close #1
End Sub

Private Sub cmdSortByCountry_Click()
'clears picture box
'sorts the top15golfers alphabetically by their country and keeps all the info on that line in order
picTop15.Cls
    For Pass = 1 To Counter - 1
        For Pos = 1 To Counter - Pass
            If Country(Pos) > Country(Pos + 1) Then
                TempGolferName = GolferName(Pos)
                GolferName(Pos) = GolferName(Pos + 1)
                GolferName(Pos + 1) = TempGolferName
                
                'each variable parralel to the golfers country needs to be sorted to stay with the corresponding data
                TempCountry = Country(Pos)
                Country(Pos) = Country(Pos + 1)
                Country(Pos + 1) = TempCountry
                
                TempPGATourWins = PGATourWins(Pos)
                PGATourWins(Pos) = PGATourWins(Pos + 1)
                PGATourWins(Pos + 1) = TempPGATourWins
                
                TempYearTurnedPro = YearTurnedPro(Pos)
                YearTurnedPro(Pos) = YearTurnedPro(Pos + 1)
                YearTurnedPro(Pos + 1) = TempYearTurnedPro
                
                TempRank = Rank(Pos)
                Rank(Pos) = Rank(Pos + 1)
                Rank(Pos + 1) = TempRank
            End If
        Next Pos
    Next Pass
'prints the heading for the table
picTop15.Print "Top 15 Professional Golfers Sorted By Their Country:"
picTop15.Print
picTop15.Print "Rank"; Tab(10); "Golfer Name"; Tab(40); "Country"; Tab(65); "Year Turned Pro"; Tab(95); "PGA Tour Wins"
picTop15.Print "_________________________________________________________________________________________________________"
'prints the sorted data into the picture box in alphabetical order by country
For I = 1 To Counter
    picTop15.Print Tab(2); Rank(I); Tab(10); GolferName(I); Tab(40); Country(I); Tab(69); YearTurnedPro(I); Tab(100); PGATourWins(I)
Next I
End Sub

Private Sub cmdSortByName_Click()
'clears picture box
'sorts the top15golfers alphbetically by their name and keeps all the info on that line in order
    picTop15.Cls
    For Pass = 1 To Counter - 1
        For Pos = 1 To Counter - Pass
            If GolferName(Pos) > GolferName(Pos + 1) Then
                
                TempGolferName = GolferName(Pos)
                GolferName(Pos) = GolferName(Pos + 1)
                GolferName(Pos + 1) = TempGolferName
                
                'each variable parralel to the golfers name needs to be sorted to stay with the corresponding data
                TempCountry = Country(Pos)
                Country(Pos) = Country(Pos + 1)
                Country(Pos + 1) = TempCountry
                
                TempPGATourWins = PGATourWins(Pos)
                PGATourWins(Pos) = PGATourWins(Pos + 1)
                PGATourWins(Pos + 1) = TempPGATourWins
                
                TempYearTurnedPro = YearTurnedPro(Pos)
                YearTurnedPro(Pos) = YearTurnedPro(Pos + 1)
                YearTurnedPro(Pos + 1) = TempYearTurnedPro
                
                TempRank = Rank(Pos)
                Rank(Pos) = Rank(Pos + 1)
                Rank(Pos + 1) = TempRank
            End If
        Next Pos
    Next Pass
'prints the heading of the table
picTop15.Print "Top 15 Professional Golfers Sorted By Their Name:"
picTop15.Print
picTop15.Print "Rank"; Tab(10); "Golfer Name"; Tab(40); "Country"; Tab(65); "Year Turned Pro"; Tab(95); "PGA Tour Wins"
picTop15.Print "_________________________________________________________________________________________________________"
'prints the sorted data in alphabetical order by the golfers name
For I = 1 To Counter
    picTop15.Print Tab(2); Rank(I); Tab(10); GolferName(I); Tab(40); Country(I); Tab(69); YearTurnedPro(I); Tab(100); PGATourWins(I)
Next I
End Sub

Private Sub cmdSortByWins_Click()
'clears the picture box
'sorts the top15golfers in decending order by their number of wins and keeps all the info on that line in order
    picTop15.Cls
    For Pass = 1 To Counter - 1
        For Pos = 1 To Counter - Pass
            If PGATourWins(Pos) < PGATourWins(Pos + 1) Then
                TempGolferName = GolferName(Pos)
                GolferName(Pos) = GolferName(Pos + 1)
                GolferName(Pos + 1) = TempGolferName
                
                'each variable parralel to the golfers number of wins needs to be sorted to stay with the corresponding data
                TempCountry = Country(Pos)
                Country(Pos) = Country(Pos + 1)
                Country(Pos + 1) = TempCountry
                
                TempPGATourWins = PGATourWins(Pos)
                PGATourWins(Pos) = PGATourWins(Pos + 1)
                PGATourWins(Pos + 1) = TempPGATourWins
                
                TempYearTurnedPro = YearTurnedPro(Pos)
                YearTurnedPro(Pos) = YearTurnedPro(Pos + 1)
                YearTurnedPro(Pos + 1) = TempYearTurnedPro
                
                TempRank = Rank(Pos)
                Rank(Pos) = Rank(Pos + 1)
                Rank(Pos + 1) = TempRank
            End If
        Next Pos
    Next Pass
'prints the heading and labels for the table
picTop15.Print "Top 15 Professional Golfers Sorted By Their Career Number of PGA Tour Wins:"
picTop15.Print
picTop15.Print "Rank"; Tab(10); "Golfer Name"; Tab(40); "Country"; Tab(65); "Year Turned Pro"; Tab(95); "PGA Tour Wins"
picTop15.Print "_________________________________________________________________________________________________________"
'prints the contents of the arrays in descending order by their number of wins
For I = 1 To Counter
    picTop15.Print Tab(2); Rank(I); Tab(10); GolferName(I); Tab(40); Country(I); Tab(69); YearTurnedPro(I); Tab(100); PGATourWins(I)
Next I
End Sub

Private Sub cmdSortByYearTurnedPro_Click()
    picTop15.Cls
'clears picture box
'sorts the top15golfers by their year turned pro in ascending order and keeps all the info on that line in order
    For Pass = 1 To Counter - 1
        For Pos = 1 To Counter - Pass
            If YearTurnedPro(Pos) > YearTurnedPro(Pos + 1) Then
                TempGolferName = GolferName(Pos)
                GolferName(Pos) = GolferName(Pos + 1)
                GolferName(Pos + 1) = TempGolferName
                
                'each variable parralel to the golfers year turned pro needs to be sorted to stay with the corresponding data
                TempCountry = Country(Pos)
                Country(Pos) = Country(Pos + 1)
                Country(Pos + 1) = TempCountry
                
                TempPGATourWins = PGATourWins(Pos)
                PGATourWins(Pos) = PGATourWins(Pos + 1)
                PGATourWins(Pos + 1) = TempPGATourWins
                
                TempYearTurnedPro = YearTurnedPro(Pos)
                YearTurnedPro(Pos) = YearTurnedPro(Pos + 1)
                YearTurnedPro(Pos + 1) = TempYearTurnedPro
                
                TempRank = Rank(Pos)
                Rank(Pos) = Rank(Pos + 1)
                Rank(Pos + 1) = TempRank
            End If
        Next Pos
    Next Pass
'Prints the heading and the labels for the table
picTop15.Print "Top 15 Professional Golfers Sorted By The Year That He Turned Pro:"
picTop15.Print
picTop15.Print "Rank"; Tab(10); "Golfer Name"; Tab(40); "Country"; Tab(65); "Year Turned Pro"; Tab(95); "PGA Tour Wins"
picTop15.Print "_________________________________________________________________________________________________________"
'prints the data from the arrays sorted in ascending order by the year they turned pro
For I = 1 To Counter
    picTop15.Print Tab(2); Rank(I); Tab(10); GolferName(I); Tab(40); Country(I); Tab(69); YearTurnedPro(I); Tab(100); PGATourWins(I)
Next I
End Sub
