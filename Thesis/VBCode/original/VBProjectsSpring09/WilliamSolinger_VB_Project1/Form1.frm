VERSION 5.00
Begin VB.Form frmSort 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   13260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFormCalculations 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Go To the Calculations Form"
      Height          =   1695
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdFormSearch 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Go To the Searching Form"
      Height          =   1695
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortI 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort By Position"
      Height          =   975
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortH 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort By Last Name"
      Height          =   975
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   2295
   End
   Begin VB.CommandButton cmdSortG 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort By Average"
      Height          =   975
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSortF 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort By Slugging Percentage"
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSortE 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort By On Base Percentage"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSortD 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort By RBIs"
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSortC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort By Home Runs"
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSortB 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort By Hits"
      Height          =   975
      Left            =   5640
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSortA 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort By At Bats"
      Height          =   3015
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1455
   End
   Begin VB.PictureBox picTwinsTerritory 
      Height          =   1575
      Left            =   1920
      ScaleHeight     =   1515
      ScaleWidth      =   8715
      TabIndex        =   3
      Top             =   3120
      Width           =   8775
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      Height          =   1575
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      Height          =   2655
      Left            =   1920
      ScaleHeight     =   2595
      ScaleWidth      =   10635
      TabIndex        =   1
      Top             =   360
      Width           =   10695
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Read From File and Show Players in Batting Order"
      Height          =   4335
      Left            =   360
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2008 Minnesota Twins
'Sorting the Twins by Statistics
'Bill Solinger
'March 24, 2009
'This project will read in a file with the 2008 Minnesota Twins starting roster.
'The user will be able to sort by statistic, search for statistics,
'see pictures of each player, and see how the calculations for some
'of the statistics are done.
'This form in particular will first read in the 2008 Minnesota Twins starting roster.
'Then, the user will be able to sort the roster according to each statistic.

Private Sub cmdFormCalculations_Click()
     'This button will bring the user to the Calculations form.
     frmSort.Visible = False
    frmCalculations.Visible = True
End Sub

Private Sub cmdFormSearch_Click()
    'This button will bring the user to the Searching form.
    frmSort.Visible = False
    frmSearch.Visible = True
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRead_Click()
    'This button first reads in a file from Notepad, and then stores the data into different arrays.  Then, it prints each player and all his statistics.
    Open App.Path & "\Twins.txt" For Input As #1
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, PlayerName(CTR), Position(CTR), AtBats(CTR), Hits(CTR), HomeRuns(CTR), RBI(CTR), OnBasePercentage(CTR), SluggingPercentage(CTR), Average(CTR)
        picResults.Print PlayerName(CTR); Tab(19); Position(CTR); Tab(38); AtBats(CTR), Hits(CTR), HomeRuns(CTR), RBI(CTR), FormatNumber(OnBasePercentage(CTR), 3), FormatNumber(SluggingPercentage(CTR), 3), FormatNumber(Average(CTR), 3)
    Loop
    Close #1
End Sub

Private Sub cmdSortA_Click()
    'This button will sort the players by At Bats into descending order.
    
    Dim Pass As Integer, Pos As Integer, i As Integer
    Dim tempName As String, tempPosition As String, tempAtBats As Integer, tempHits As Integer
    Dim tempHR As Integer, tempRBI As Integer, tempOBP As Single, tempSLG As Single, tempAVG As Single
    
    
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If AtBats(Pos) < AtBats(Pos + 1) Then
                tempName = PlayerName(Pos)
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = tempName
                tempPosition = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = tempPosition
                tempAtBats = AtBats(Pos)
                AtBats(Pos) = AtBats(Pos + 1)
                AtBats(Pos + 1) = tempAtBats
                tempHits = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = tempHits
                tempHR = HomeRuns(Pos)
                HomeRuns(Pos) = HomeRuns(Pos + 1)
                HomeRuns(Pos + 1) = tempHR
                tempRBI = RBI(Pos)
                RBI(Pos) = RBI(Pos + 1)
                RBI(Pos + 1) = tempRBI
                tempOBP = OnBasePercentage(Pos)
                OnBasePercentage(Pos) = OnBasePercentage(Pos + 1)
                OnBasePercentage(Pos + 1) = tempOBP
                tempSLG = SluggingPercentage(Pos)
                SluggingPercentage(Pos) = SluggingPercentage(Pos + 1)
                SluggingPercentage(Pos + 1) = tempSLG
                tempAVG = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAVG
            End If
        Next Pos
    Next Pass
    For i = 1 To CTR
        picResults.Print PlayerName(i); Tab(19); Position(i); Tab(38); AtBats(i), Hits(i), HomeRuns(i), RBI(i), FormatNumber(OnBasePercentage(i), 3), FormatNumber(SluggingPercentage(i), 3), FormatNumber(Average(i), 3)
    Next i
End Sub

Private Sub cmdSortB_Click()
    'This button will sort the players by Hits into descending order.
    Dim Pass As Integer, Pos As Integer, i As Integer
    Dim tempName As String, tempPosition As String, tempAtBats As Integer, tempHits As Integer
    Dim tempHR As Integer, tempRBI As Integer, tempOBP As Single, tempSLG As Single, tempAVG As Single
    
    
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Hits(Pos) < Hits(Pos + 1) Then
                tempName = PlayerName(Pos)
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = tempName
                tempPosition = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = tempPosition
                tempAtBats = AtBats(Pos)
                AtBats(Pos) = AtBats(Pos + 1)
                AtBats(Pos + 1) = tempAtBats
                tempHits = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = tempHits
                tempHR = HomeRuns(Pos)
                HomeRuns(Pos) = HomeRuns(Pos + 1)
                HomeRuns(Pos + 1) = tempHR
                tempRBI = RBI(Pos)
                RBI(Pos) = RBI(Pos + 1)
                RBI(Pos + 1) = tempRBI
                tempOBP = OnBasePercentage(Pos)
                OnBasePercentage(Pos) = OnBasePercentage(Pos + 1)
                OnBasePercentage(Pos + 1) = tempOBP
                tempSLG = SluggingPercentage(Pos)
                SluggingPercentage(Pos) = SluggingPercentage(Pos + 1)
                SluggingPercentage(Pos + 1) = tempSLG
                tempAVG = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAVG
            End If
        Next Pos
    Next Pass
    For i = 1 To CTR
        picResults.Print PlayerName(i); Tab(19); Position(i); Tab(38); AtBats(i), Hits(i), HomeRuns(i), RBI(i), FormatNumber(OnBasePercentage(i), 3), FormatNumber(SluggingPercentage(i), 3), FormatNumber(Average(i), 3)
    Next i
End Sub

Private Sub cmdSortC_Click()
    'This button will sort the players by Home Runs into descending order.
    Dim Pass As Integer, Pos As Integer, i As Integer
    Dim tempName As String, tempPosition As String, tempAtBats As Integer, tempHits As Integer
    Dim tempHR As Integer, tempRBI As Integer, tempOBP As Single, tempSLG As Single, tempAVG As Single
    
    
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If HomeRuns(Pos) < HomeRuns(Pos + 1) Then
                tempName = PlayerName(Pos)
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = tempName
                tempPosition = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = tempPosition
                tempAtBats = AtBats(Pos)
                AtBats(Pos) = AtBats(Pos + 1)
                AtBats(Pos + 1) = tempAtBats
                tempHits = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = tempHits
                tempHR = HomeRuns(Pos)
                HomeRuns(Pos) = HomeRuns(Pos + 1)
                HomeRuns(Pos + 1) = tempHR
                tempRBI = RBI(Pos)
                RBI(Pos) = RBI(Pos + 1)
                RBI(Pos + 1) = tempRBI
                tempOBP = OnBasePercentage(Pos)
                OnBasePercentage(Pos) = OnBasePercentage(Pos + 1)
                OnBasePercentage(Pos + 1) = tempOBP
                tempSLG = SluggingPercentage(Pos)
                SluggingPercentage(Pos) = SluggingPercentage(Pos + 1)
                SluggingPercentage(Pos + 1) = tempSLG
                tempAVG = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAVG
            End If
        Next Pos
    Next Pass
    For i = 1 To CTR
        picResults.Print PlayerName(i); Tab(19); Position(i); Tab(38); AtBats(i), Hits(i), HomeRuns(i), RBI(i), FormatNumber(OnBasePercentage(i), 3), FormatNumber(SluggingPercentage(i), 3), FormatNumber(Average(i), 3)
    Next i
End Sub

Private Sub cmdSortD_Click()
    'This button will sort the players by RBI into descending order.
    Dim Pass As Integer, Pos As Integer, i As Integer
    Dim tempName As String, tempPosition As String, tempAtBats As Integer, tempHits As Integer
    Dim tempHR As Integer, tempRBI As Integer, tempOBP As Single, tempSLG As Single, tempAVG As Single
    
    
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If RBI(Pos) < RBI(Pos + 1) Then
                tempName = PlayerName(Pos)
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = tempName
                tempPosition = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = tempPosition
                tempAtBats = AtBats(Pos)
                AtBats(Pos) = AtBats(Pos + 1)
                AtBats(Pos + 1) = tempAtBats
                tempHits = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = tempHits
                tempHR = HomeRuns(Pos)
                HomeRuns(Pos) = HomeRuns(Pos + 1)
                HomeRuns(Pos + 1) = tempHR
                tempRBI = RBI(Pos)
                RBI(Pos) = RBI(Pos + 1)
                RBI(Pos + 1) = tempRBI
                tempOBP = OnBasePercentage(Pos)
                OnBasePercentage(Pos) = OnBasePercentage(Pos + 1)
                OnBasePercentage(Pos + 1) = tempOBP
                tempSLG = SluggingPercentage(Pos)
                SluggingPercentage(Pos) = SluggingPercentage(Pos + 1)
                SluggingPercentage(Pos + 1) = tempSLG
                tempAVG = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAVG
            End If
        Next Pos
    Next Pass
    For i = 1 To CTR
        picResults.Print PlayerName(i); Tab(19); Position(i); Tab(38); AtBats(i), Hits(i), HomeRuns(i), RBI(i), FormatNumber(OnBasePercentage(i), 3), FormatNumber(SluggingPercentage(i), 3), FormatNumber(Average(i), 3)
    Next i
End Sub

Private Sub cmdSortE_Click()
    'This button will sort the players by On Base Percentage into descending order.

    Dim Pass As Integer, Pos As Integer, i As Integer
    Dim tempName As String, tempPosition As String, tempAtBats As Integer, tempHits As Integer
    Dim tempHR As Integer, tempRBI As Integer, tempOBP As Single, tempSLG As Single, tempAVG As Single
    
    
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If OnBasePercentage(Pos) < OnBasePercentage(Pos + 1) Then
                tempName = PlayerName(Pos)
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = tempName
                tempPosition = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = tempPosition
                tempAtBats = AtBats(Pos)
                AtBats(Pos) = AtBats(Pos + 1)
                AtBats(Pos + 1) = tempAtBats
                tempHits = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = tempHits
                tempHR = HomeRuns(Pos)
                HomeRuns(Pos) = HomeRuns(Pos + 1)
                HomeRuns(Pos + 1) = tempHR
                tempRBI = RBI(Pos)
                RBI(Pos) = RBI(Pos + 1)
                RBI(Pos + 1) = tempRBI
                tempOBP = OnBasePercentage(Pos)
                OnBasePercentage(Pos) = OnBasePercentage(Pos + 1)
                OnBasePercentage(Pos + 1) = tempOBP
                tempSLG = SluggingPercentage(Pos)
                SluggingPercentage(Pos) = SluggingPercentage(Pos + 1)
                SluggingPercentage(Pos + 1) = tempSLG
                tempAVG = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAVG
            End If
        Next Pos
    Next Pass
    For i = 1 To CTR
        picResults.Print PlayerName(i); Tab(19); Position(i); Tab(38); AtBats(i), Hits(i), HomeRuns(i), RBI(i), FormatNumber(OnBasePercentage(i), 3), FormatNumber(SluggingPercentage(i), 3), FormatNumber(Average(i), 3)
    Next i
End Sub

Private Sub cmdSortF_Click()

    'This button will sort the players by Slugging Percentage into descending order.
    
    Dim Pass As Integer, Pos As Integer, i As Integer
    Dim tempName As String, tempPosition As String, tempAtBats As Integer, tempHits As Integer
    Dim tempHR As Integer, tempRBI As Integer, tempOBP As Single, tempSLG As Single, tempAVG As Single
    
    
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If SluggingPercentage(Pos) < SluggingPercentage(Pos + 1) Then
                tempName = PlayerName(Pos)
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = tempName
                tempPosition = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = tempPosition
                tempAtBats = AtBats(Pos)
                AtBats(Pos) = AtBats(Pos + 1)
                AtBats(Pos + 1) = tempAtBats
                tempHits = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = tempHits
                tempHR = HomeRuns(Pos)
                HomeRuns(Pos) = HomeRuns(Pos + 1)
                HomeRuns(Pos + 1) = tempHR
                tempRBI = RBI(Pos)
                RBI(Pos) = RBI(Pos + 1)
                RBI(Pos + 1) = tempRBI
                tempOBP = OnBasePercentage(Pos)
                OnBasePercentage(Pos) = OnBasePercentage(Pos + 1)
                OnBasePercentage(Pos + 1) = tempOBP
                tempSLG = SluggingPercentage(Pos)
                SluggingPercentage(Pos) = SluggingPercentage(Pos + 1)
                SluggingPercentage(Pos + 1) = tempSLG
                tempAVG = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAVG
            End If
        Next Pos
    Next Pass
    For i = 1 To CTR
        picResults.Print PlayerName(i); Tab(19); Position(i); Tab(38); AtBats(i), Hits(i), HomeRuns(i), RBI(i), FormatNumber(OnBasePercentage(i), 3), FormatNumber(SluggingPercentage(i), 3), FormatNumber(Average(i), 3)
    Next i
End Sub

Private Sub cmdSortG_Click()
    'This button will sort the players by Batting Average into descending order.

    Dim Pass As Integer, Pos As Integer, i As Integer
    Dim tempName As String, tempPosition As String, tempAtBats As Integer, tempHits As Integer
    Dim tempHR As Integer, tempRBI As Integer, tempOBP As Single, tempSLG As Single, tempAVG As Single
    
    
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Average(Pos) < Average(Pos + 1) Then
                tempName = PlayerName(Pos)
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = tempName
                tempPosition = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = tempPosition
                tempAtBats = AtBats(Pos)
                AtBats(Pos) = AtBats(Pos + 1)
                AtBats(Pos + 1) = tempAtBats
                tempHits = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = tempHits
                tempHR = HomeRuns(Pos)
                HomeRuns(Pos) = HomeRuns(Pos + 1)
                HomeRuns(Pos + 1) = tempHR
                tempRBI = RBI(Pos)
                RBI(Pos) = RBI(Pos + 1)
                RBI(Pos + 1) = tempRBI
                tempOBP = OnBasePercentage(Pos)
                OnBasePercentage(Pos) = OnBasePercentage(Pos + 1)
                OnBasePercentage(Pos + 1) = tempOBP
                tempSLG = SluggingPercentage(Pos)
                SluggingPercentage(Pos) = SluggingPercentage(Pos + 1)
                SluggingPercentage(Pos + 1) = tempSLG
                tempAVG = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAVG
            End If
        Next Pos
    Next Pass
    For i = 1 To CTR
        picResults.Print PlayerName(i); Tab(19); Position(i); Tab(38); AtBats(i), Hits(i), HomeRuns(i), RBI(i), FormatNumber(OnBasePercentage(i), 3), FormatNumber(SluggingPercentage(i), 3), FormatNumber(Average(i), 3)
    Next i
End Sub

Private Sub cmdSortH_Click()
    'This button will sort the players into alphabetical order by last name.

    Dim Pass As Integer, Pos As Integer, i As Integer
    Dim tempName As String, tempPosition As String, tempAtBats As Integer, tempHits As Integer
    Dim tempHR As Integer, tempRBI As Integer, tempOBP As Single, tempSLG As Single, tempAVG As Single
    
    
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If PlayerName(Pos) > PlayerName(Pos + 1) Then
                tempName = PlayerName(Pos)
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = tempName
                tempPosition = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = tempPosition
                tempAtBats = AtBats(Pos)
                AtBats(Pos) = AtBats(Pos + 1)
                AtBats(Pos + 1) = tempAtBats
                tempHits = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = tempHits
                tempHR = HomeRuns(Pos)
                HomeRuns(Pos) = HomeRuns(Pos + 1)
                HomeRuns(Pos + 1) = tempHR
                tempRBI = RBI(Pos)
                RBI(Pos) = RBI(Pos + 1)
                RBI(Pos + 1) = tempRBI
                tempOBP = OnBasePercentage(Pos)
                OnBasePercentage(Pos) = OnBasePercentage(Pos + 1)
                OnBasePercentage(Pos + 1) = tempOBP
                tempSLG = SluggingPercentage(Pos)
                SluggingPercentage(Pos) = SluggingPercentage(Pos + 1)
                SluggingPercentage(Pos + 1) = tempSLG
                tempAVG = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAVG
            End If
        Next Pos
    Next Pass
    For i = 1 To CTR
        picResults.Print PlayerName(i); Tab(19); Position(i); Tab(38); AtBats(i), Hits(i), HomeRuns(i), RBI(i), FormatNumber(OnBasePercentage(i), 3), FormatNumber(SluggingPercentage(i), 3), FormatNumber(Average(i), 3)
    Next i
End Sub

Private Sub cmdSortI_Click()
    'This button will sort the players by position in alphabetical order.
    Dim Pass As Integer, Pos As Integer, i As Integer
    Dim tempName As String, tempPosition As String, tempAtBats As Integer, tempHits As Integer
    Dim tempHR As Integer, tempRBI As Integer, tempOBP As Single, tempSLG As Single, tempAVG As Single
    
    
    picResults.Cls
    picResults.Print "2008 Minnesota Twins Starting Order and Statistics"
    picResults.Print "******************************************************************************************************************************************************************************"
    picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
    picResults.Print "******************************************************************************************************************************************************************************"
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Position(Pos) > Position(Pos + 1) Then
                tempName = PlayerName(Pos)
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = tempName
                tempPosition = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = tempPosition
                tempAtBats = AtBats(Pos)
                AtBats(Pos) = AtBats(Pos + 1)
                AtBats(Pos + 1) = tempAtBats
                tempHits = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = tempHits
                tempHR = HomeRuns(Pos)
                HomeRuns(Pos) = HomeRuns(Pos + 1)
                HomeRuns(Pos + 1) = tempHR
                tempRBI = RBI(Pos)
                RBI(Pos) = RBI(Pos + 1)
                RBI(Pos + 1) = tempRBI
                tempOBP = OnBasePercentage(Pos)
                OnBasePercentage(Pos) = OnBasePercentage(Pos + 1)
                OnBasePercentage(Pos + 1) = tempOBP
                tempSLG = SluggingPercentage(Pos)
                SluggingPercentage(Pos) = SluggingPercentage(Pos + 1)
                SluggingPercentage(Pos + 1) = tempSLG
                tempAVG = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAVG
            End If
        Next Pos
    Next Pass
    For i = 1 To CTR
        picResults.Print PlayerName(i); Tab(19); Position(i); Tab(38); AtBats(i), Hits(i), HomeRuns(i), RBI(i), FormatNumber(OnBasePercentage(i), 3), FormatNumber(SluggingPercentage(i), 3), FormatNumber(Average(i), 3)
    Next i
End Sub

Private Sub Form_Load()
    'This picture automatically loads when the form is opened.
    picTwinsTerritory.Picture = LoadPicture(App.Path & "\TwinsTerritory.jpg")
End Sub
