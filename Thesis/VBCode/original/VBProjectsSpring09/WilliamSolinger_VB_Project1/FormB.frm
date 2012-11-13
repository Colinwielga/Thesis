VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   9990
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFormCalculations 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go To the Calculations Form"
      Height          =   1335
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton cmdShowG 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show All Players with a Higher Average than Your Input Number"
      Height          =   1095
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8520
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowF 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show All Players with a Higher Slugging Percentage than Your Input Number"
      Height          =   1095
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowE 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show All Players with a Higher On Base Percentage than Your Input Number"
      Height          =   1095
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowD 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show All Players with More RBI than Your Input Number"
      Height          =   1095
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show All Players with More Home Runs than Your Input Number"
      Height          =   1095
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowB 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show All Players with More Hits than Your Input Number"
      Height          =   1095
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowA 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show All Players with More At Bats than Your Input Number"
      Height          =   1095
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   2895
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   13080
      ScaleHeight     =   1875
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowPlayer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type in a player's last name, and see their picture."
      Height          =   1935
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   4455
   End
   Begin VB.PictureBox picMTwins 
      BackColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   240
      ScaleHeight     =   6795
      ScaleWidth      =   8115
      TabIndex        =   4
      Top             =   2880
      Width           =   8175
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Players in Batting Order"
      Height          =   2655
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   2160
      ScaleHeight     =   2595
      ScaleWidth      =   10395
      TabIndex        =   2
      Top             =   120
      Width           =   10455
   End
   Begin VB.CommandButton cmdFormSort 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go To the Sorting Form"
      Height          =   1335
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   2655
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2008 Minnesota Twins
'Search for Statistics and Show Me the Players!
'Bill Solinger
'March 24, 2009
'This form will allow the user to give a number for a certain statistic,
'using an input box, and the program will show all players with a
'higher number in that specific statistic.
'Also, the user will input a player's name and be able to see his picture.

Private Sub cmdFormCalculations_Click()
    'This button will bring the user to the Calculations form.
    frmSearch.Visible = False
    frmCalculations.Visible = True
End Sub

Private Sub cmdFormSort_Click()
    'This button will bring the user to the Sorting form.
    frmSort.Visible = True
    frmSearch.Visible = False
End Sub

Private Sub cmdPrint_Click()
    'This button will again read the Twins in from a Notepad file and list them in their batting order in the results box.
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

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdShowA_Click()
    'The user will input a value in an Input box, and the results box will show all players with over that value in At Bats.

    Dim Found As Boolean
    Dim k As Integer
    Dim SearchAtBats As Integer
    SearchAtBats = InputBox("Enter a number of At Bats so that you can see all players who had more At Bats than your number.", "At Bats")
    Found = False
       
    If SearchAtBats <= 0 Then
        MsgBox "Please enter an amount of At Bats greater than 0."
        Found = True
        
    Else:   picResults.Cls
    
            picResults.Print "2008 Minnesota Twins Players with over " & SearchAtBats & " At Bats"
            picResults.Print "******************************************************************************************************************************************************************************"
            picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
            picResults.Print "******************************************************************************************************************************************************************************"
            
            For k = 1 To CTR
                If SearchAtBats <= AtBats(k) Then
                    Found = True
                    picResults.Print PlayerName(k); Tab(19); Position(k); Tab(38); AtBats(k), Hits(k), HomeRuns(k), RBI(k), FormatNumber(OnBasePercentage(k), 3), FormatNumber(SluggingPercentage(k), 3), FormatNumber(Average(k), 3)
                End If
            Next k
    End If
    
    If (Not Found) Then
        MsgBox "There are no players with over " & SearchAtBats & " At Bats."
    End If
    
End Sub

Private Sub cmdShowB_Click()
    'The user will input a value in an Input box, and the results box will show all players with over that value in Hits.

    Dim Found As Boolean
    Dim k As Integer
    Dim SearchHits As Integer
    SearchHits = InputBox("Enter a number of Hits so that you can see all players who had more Hits than your number.", "Hits")
    Found = False
       
    If SearchHits <= 0 Then
        MsgBox "Please enter an amount of Hits greater than 0."
        Found = True
        
    Else:   picResults.Cls
    
            picResults.Print "2008 Minnesota Twins Players with over " & SearchHits & " Hits"
            picResults.Print "******************************************************************************************************************************************************************************"
            picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
            picResults.Print "******************************************************************************************************************************************************************************"
            
            For k = 1 To CTR
                If SearchHits <= Hits(k) Then
                    Found = True
                    picResults.Print PlayerName(k); Tab(19); Position(k); Tab(38); AtBats(k), Hits(k), HomeRuns(k), RBI(k), FormatNumber(OnBasePercentage(k), 3), FormatNumber(SluggingPercentage(k), 3), FormatNumber(Average(k), 3)
                End If
            Next k
    End If
    
    If (Not Found) Then
        MsgBox "There are no players with over " & SearchHits & " Hits."
    End If
End Sub

Private Sub cmdShowC_Click()
    'The user will input a value in an Input box, and the results box will show all players with over that value in Home Runs.
    
    Dim Found As Boolean
    Dim k As Integer
    Dim SearchHR As Integer
    SearchHR = InputBox("Enter a number of Home Runs so that you can see all players who had more Home Runs than your number.", "Home Runs")
    Found = False
       
    If SearchHR <= 0 Then
        MsgBox "Please enter a Home Run amount greater than 0."
        Found = True
        
    Else:   picResults.Cls
    
            picResults.Print "2008 Minnesota Twins Players with over " & SearchHR & " Home Runs"
            picResults.Print "******************************************************************************************************************************************************************************"
            picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
            picResults.Print "******************************************************************************************************************************************************************************"
            
            For k = 1 To CTR
                If SearchHR <= HomeRuns(k) Then
                    Found = True
                    picResults.Print PlayerName(k); Tab(19); Position(k); Tab(38); AtBats(k), Hits(k), HomeRuns(k), RBI(k), FormatNumber(OnBasePercentage(k), 3), FormatNumber(SluggingPercentage(k), 3), FormatNumber(Average(k), 3)
                End If
            Next k
    End If
    
    If (Not Found) Then
        MsgBox "There are no players with over " & SearchHR & " Home Runs."
    End If
End Sub

Private Sub cmdShowD_Click()
    'The user will input a value in an Input box, and the results box will show all players with over that value in RBI.
    
    Dim Found As Boolean
    Dim k As Integer
    Dim SearchRBI As Integer
    SearchRBI = InputBox("Enter a number of Runs Batted In so that you can see all players who had more Runs Batted In than your number.", "RBI")
    Found = False
   
    If SearchRBI <= 0 Then
        MsgBox "Please enter an RBI amount greater than 0."
        Found = True
        
    Else:   picResults.Cls
    
            picResults.Print "2008 Minnesota Twins Players with over " & SearchRBI & " Runs Batted In"
            picResults.Print "******************************************************************************************************************************************************************************"
            picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
            picResults.Print "******************************************************************************************************************************************************************************"
            
            For k = 1 To CTR
                If SearchRBI <= RBI(k) Then
                    Found = True
                    picResults.Print PlayerName(k); Tab(19); Position(k); Tab(38); AtBats(k), Hits(k), HomeRuns(k), RBI(k), FormatNumber(OnBasePercentage(k), 3), FormatNumber(SluggingPercentage(k), 3), FormatNumber(Average(k), 3)
                End If
            Next k
    End If
    
    If (Not Found) Then
        MsgBox "There are no players with over " & SearchRBI & " Runs Batted In."
    End If
End Sub

Private Sub cmdShowE_Click()
    'The user will input a value in an Input box, and the results box will show all players with over that value in On Base Percentage.
    
    Dim Found As Boolean
    Dim k As Integer
    Dim SearchOBP As Single
    SearchOBP = InputBox("Enter an On Base Percentage (from 0.000 to 1.000) so that you can see all players who had a higher On Base Percentage than your number.", "OBP")
    Found = False
   
    If SearchOBP <= 0 Or SearchOBP >= 1 Then
        MsgBox "Please enter an On Base Percentage between .000 and 1.000."
        Found = True
        
    Else:       picResults.Cls
    
                picResults.Print "2008 Minnesota Twins Players with over a " & FormatNumber(SearchOBP, 3) & " On Base Percentage"
                picResults.Print "******************************************************************************************************************************************************************************"
                picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
                picResults.Print "******************************************************************************************************************************************************************************"
    
                For k = 1 To CTR
                    If SearchOBP <= OnBasePercentage(k) Then
                        Found = True
                        picResults.Print PlayerName(k); Tab(19); Position(k); Tab(38); AtBats(k), Hits(k), HomeRuns(k), RBI(k), FormatNumber(OnBasePercentage(k), 3), FormatNumber(SluggingPercentage(k), 3), FormatNumber(Average(k), 3)
                    End If
                Next k
    End If
    
    If (Not Found) Then
        MsgBox "There are no players with over a " & SearchOBP & " On Base Percentage."
    End If
End Sub

Private Sub cmdShowF_Click()
    'The user will input a value in an Input box, and the results box will show all players with over that value in Slugging Percentage.
    
    Dim Found As Boolean
    Dim k As Integer
    Dim SearchSLG As Single
    SearchSLG = InputBox("Enter a Slugging Percentage so that you can see all players who had a higher Slugging Percentage than your number.", "SLG")
    
    Found = False
   
    If SearchSLG <= 0 Then
        MsgBox "Please enter a Slugging Percentage above 0."
        Found = True
        
    Else:   picResults.Cls
    
            picResults.Print "2008 Minnesota Twins Players with over a " & FormatNumber(SearchSLG, 3) & " Slugging Percentage"
            picResults.Print "******************************************************************************************************************************************************************************"
            picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
            picResults.Print "******************************************************************************************************************************************************************************"
    
            For k = 1 To CTR
                If SearchSLG <= SluggingPercentage(k) Then
                    Found = True
                    picResults.Print PlayerName(k); Tab(19); Position(k); Tab(38); AtBats(k), Hits(k), HomeRuns(k), RBI(k), FormatNumber(OnBasePercentage(k), 3), FormatNumber(SluggingPercentage(k), 3), FormatNumber(Average(k), 3)
                End If
            Next k
    End If
    
    If (Not Found) Then
        MsgBox "There are no players with over a " & SearchSLG & " Slugging Percentage."
    End If
End Sub

Private Sub cmdShowG_Click()
    ''The user will input a value in an Input box, and the results box will show all players with over that value in Average.
    
    Dim Found As Boolean
    Dim k As Integer
    Dim SearchAVG As Single
    SearchAVG = InputBox("Enter an Average (from .000 to 1.000) so that you can see all players who had a higher Average than your number.", "AVG")

    Found = False
   
    If SearchAVG <= 0 Or SearchAVG >= 1 Then
        MsgBox "Please enter an Average between .000 and 1.000."
    
    Else:   picResults.Cls
    
            picResults.Print "2008 Minnesota Twins Players with over a " & FormatNumber(SearchAVG, 3) & " Average"
            picResults.Print "******************************************************************************************************************************************************************************"
            picResults.Print "Name"; Tab(19); "Position"; Tab(38); "At Bats", "Hits", "Home Runs", "RBI", "OBP", "SLG", "Average"
            picResults.Print "******************************************************************************************************************************************************************************"
    
            For k = 1 To CTR
                If SearchAVG <= Average(k) Then
                    Found = True
                    picResults.Print PlayerName(k); Tab(19); Position(k); Tab(38); AtBats(k), Hits(k), HomeRuns(k), RBI(k), FormatNumber(OnBasePercentage(k), 3), FormatNumber(SluggingPercentage(k), 3), FormatNumber(Average(k), 3)
                End If
            Next k
    End If
    
    If (Not Found) Then
        MsgBox "There are no players with over a " & SearchAVG & " Average."
    End If
End Sub

Private Sub cmdShowPlayer_Click()
    'This button allows the user to type in one of the nine players' last names, and the picture box next to it will show a picture of that player.
    'The user must spell the name correctly.
    
    Dim LastName As String
    LastName = InputBox("Enter the player's last name you would like to see.", "Last Name")
    
    Select Case LastName
    Case "Span"
        picPlayer.Picture = LoadPicture(App.Path & "\Span.jpg")
    Case "Casilla"
        picPlayer.Picture = LoadPicture(App.Path & "\Casilla.jpg")
    Case "Mauer"
        picPlayer.Picture = LoadPicture(App.Path & "\Mauer.jpg")
    Case "Morneau"
        picPlayer.Picture = LoadPicture(App.Path & "\Morneau.jpg")
    Case "Young"
        picPlayer.Picture = LoadPicture(App.Path & "\Young.jpg")
    Case "Kubel"
        picPlayer.Picture = LoadPicture(App.Path & "\Kubel.jpg")
    Case "Harris"
        picPlayer.Picture = LoadPicture(App.Path & "\Harris.jpg")
    Case "Gomez"
        picPlayer.Picture = LoadPicture(App.Path & "\Gomez.jpg")
    Case "Punto"
        picPlayer.Picture = LoadPicture(App.Path & "\Punto.jpg")
    Case Else: MsgBox "Please enter a valid player from the 2008 Minnesota Twins Starting Roster.", , "Error"
    End Select
            
End Sub

Private Sub Form_Load()
    'The form automatically loads this picture.
    picMTwins.Picture = LoadPicture(App.Path & "\MTwins.jpg")
End Sub

