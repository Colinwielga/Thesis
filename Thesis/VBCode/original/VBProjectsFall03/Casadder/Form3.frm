VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF0000&
   Caption         =   "2003 Minnesota Twins Batting Statistics by Charlie Sadder"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form3"
   ScaleHeight     =   7515
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMeetB 
      Caption         =   "Individual Batter Statistics"
      Height          =   735
      Left            =   240
      TabIndex        =   26
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSb 
      Caption         =   "SB"
      Height          =   375
      Left            =   9360
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdAvg 
      Caption         =   "AVG"
      Height          =   375
      Left            =   10080
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdRbi 
      Caption         =   "RBI"
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdHr 
      Caption         =   "HR"
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdTrip 
      Caption         =   "3B"
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdDoub 
      Caption         =   "2B"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdHits 
      Caption         =   "H"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdRuns 
      Caption         =   "R"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdAtbats 
      Caption         =   "AB"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdGames 
      Caption         =   "G"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H000000FF&
      Height          =   5895
      Left            =   2280
      ScaleHeight     =   5835
      ScaleWidth      =   8595
      TabIndex        =   4
      Top             =   1560
      Width           =   8655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSwitchP 
      Caption         =   "Switch to Pitching Statistics"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdLbat 
      Caption         =   "All Batting Statistics"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Batting Average"
      Height          =   375
      Left            =   10080
      TabIndex        =   25
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stolen Bases"
      Height          =   375
      Left            =   9360
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Runs Batted In"
      Height          =   615
      Left            =   8640
      TabIndex        =   23
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Home Runs"
      Height          =   375
      Left            =   7920
      TabIndex        =   22
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Triples"
      Height          =   255
      Left            =   7320
      TabIndex        =   21
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Doubles"
      Height          =   255
      Left            =   6600
      TabIndex        =   20
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hits"
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Runs"
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "At Bats"
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.Label label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Games"
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select Category for Batting Leaders:"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1890
      Index           =   0
      Left            =   120
      Picture         =   "Form3.frx":0000
      Top             =   3360
      Width           =   1890
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MinnesotaTwinsVB (Charlie Sadder's VB project)
'Form Name : 2003 Minnesota Twins Batting Statistics (form3.frm)
'Author: Charlie Sadder
'Date Written: October 28, 2003
'Purpose of Form: To see the Twins' leaders of statistics in the 2003 season.
                  'There are different categories for batting
                  'for which you can see team leaders.

'Option Explicit is a command to force
'the user to declare all variables before the program will run.
Option Explicit
Dim i As Integer
    Dim strName(1 To 19) As String
    Dim games(1 To 19) As Integer
    Dim atbat(1 To 19) As Integer
    Dim runs(1 To 19) As Integer
    Dim hits(1 To 19) As Integer
    Dim doub(1 To 19) As Integer
    Dim trip(1 To 19) As Integer
    Dim hr(1 To 19) As Integer
    Dim rbi(1 To 19) As Integer
    Dim stolenbase(1 To 19) As Integer
    Dim avg(1 To 19) As Single

Private Sub cmdAtbats_Click()
'Orders the batters from most to least at bats.
Dim N As Single
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "AB"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If atbat(i) < atbat(i + 1) Then
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            tempor = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = tempor
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
            
            
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); atbat(i)
Next i
Close #1

End Sub

Private Sub cmdAvg_Click()
'Orders the batters from highest to lowest batting average.
Dim N As Single
Dim temp As Single
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "AVG"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If avg(i) < avg(i + 1) Then
            temp = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); avg(i)
Next i
Close #1

End Sub

Private Sub cmdClear_Click()
'Clears anything in the picture box.
    pbxResults.Cls
End Sub

Private Sub cmdDoub_Click()
'Orders the batters from most to least doubles.
Dim N As Single
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "2B"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If doub(i) < doub(i + 1) Then
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            tempor = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = tempor
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); doub(i)
Next i
Close #1

End Sub

Private Sub cmdGames_Click()
'Orders the batters from most to least games played.
Dim N As Single
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "G"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If games(i) < games(i + 1) Then
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            tempor = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = tempor
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); games(i)
Next i
Close #1
End Sub
Private Sub cmdLoadB_Click()
'Loads the list of batters and their statistics.
    pbxResults.Cls
    pbxResults.Print " Player", Tab(15); "G"; Tab(25); "W"; Tab(35); "L"; Tab(45); "SV"; Tab(55); "IP"; Tab(65); "SO"; Tab(75); "K/9"; Tab(85); "ERA"; Tab(95)
    For i = 1 To 16
        
        pbxResults.Print strName(i), Tab(15); games(i); Tab(25); wins(i); Tab(35); losses(i); Tab(45); saves(i); Tab(55); innpitch(i); Tab(65); strikeouts(i); Tab(75); ksper9(i); Tab(85); era(i); Tab(95)
    Next i
    Close #1
End Sub

Private Sub cmdHits_Click()
'Orders the batters from most to least hits.
Dim N As Single
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "H"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If hits(i) < hits(i + 1) Then
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            tempor = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = tempor
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); hits(i)
Next i
Close #1

End Sub

Private Sub cmdHr_Click()
'Orders the batters from most to least home runs.
Dim N As Single
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "HR"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If hr(i) < hr(i + 1) Then
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            tempor = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = tempor
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); hr(i)
Next i
Close #1


End Sub

Private Sub cmdLbat_Click()
'Loads the Batters and their statistics.
    pbxResults.Cls
    pbxResults.Print " Player", Tab(15); "G"; Tab(25); "AB"; Tab(35); "R"; Tab(45); "H"; Tab(55); "2B"; Tab(65); "3B"; Tab(75); "HR"; Tab(85); "RBI"; Tab(95); "SB"; Tab(105); "AVG"; Tab(115)
    For i = 1 To 19
        pbxResults.Print strName(i), Tab(15); games(i); Tab(25); atbat(i); Tab(35); runs(i); Tab(45); hits(i); Tab(55); doub(i); Tab(65); trip(i); Tab(75); hr(i); Tab(85); rbi(i); Tab(95); stolenbase(i); Tab(105); avg(i); Tab(115)
    Next i
    Close #1
    
    
End Sub



Private Sub cmdMeetB_Click()
Dim Found As Boolean
Dim iPlayer As String
Dim N As Integer
N = 16
i = 0
'has user enter a name of batter to look up in the txt
iPlayer = InputBox("Enter the name of the player you wish to find")
Found = False
'Clear whatever may be in pbxResults for repeated use.
pbxResults.Cls
'Prints header on top of the picture box
pbxResults.Print " Player", Tab(15); "G"; Tab(25); "AB"; Tab(35); "R"; Tab(45); "H"; Tab(55); "2B"; Tab(65); "3B"; Tab(75); "HR"; Tab(85); "RBI"; Tab(95); "SB"; Tab(105); "AVG"; Tab(115)
Do While i <= N - 1 And Found = False
    'counts the number of times you go through your program
    i = i + 1
    If iPlayer = strName(i) Then
    Found = True
    End If
    Loop
If Found = True Then
        'prints out the name of the batter you entered and their stats
        pbxResults.Print strName(i), Tab(15); games(i); Tab(25); atbat(i); Tab(35); runs(i); Tab(45); hits(i); Tab(55); doub(i); Tab(65); trip(i); Tab(75); hr(i); Tab(85); rbi(i); Tab(95); stolenbase(i); Tab(105); avg(i); Tab(115)
    Else
    'gives a pop up message and tells them there is no player with that name
        MsgBox ("Player not found")
End If
End Sub

Private Sub cmdRbi_Click()
'Orders the batters from most to least RBIs.
Dim N As Single
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "RBI"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If rbi(i) < rbi(i + 1) Then
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            tempor = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = tempor
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); rbi(i)
Next i
Close #1

End Sub
Private Sub cmdRuns_Click()
'Orders the batters from most to least runs scored.
Dim N As Single
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "R"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If runs(i) < runs(i + 1) Then
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            tempor = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = tempor
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); runs(i)
Next i
Close #1

End Sub

Private Sub cmdSb_Click()
'Orders the batters from most to least stolen bases.
Dim N As Single
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "SB"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If stolenbase(i) < stolenbase(i + 1) Then
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            tempor = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = tempor
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); stolenbase(i)
Next i
Close #1

End Sub

Private Sub cmdSwitchP_Click()
'Switches from form 3 to form 2.
    Form3.Hide
    Form2.Show
End Sub

Private Sub cmdQuit_Click()
'Quits the program.
    End
End Sub
Private Sub cmdTrip_Click()
'Orders the batters from most to least triples.
Dim N As Single
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 19
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "3B"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If trip(i) < trip(i + 1) Then
            temp = trip(i + 1)
            trip(i + 1) = trip(i)
            trip(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            tempor = avg(i + 1)
            avg(i + 1) = avg(i)
            avg(i) = tempor
            temp = runs(i + 1)
            runs(i + 1) = runs(i)
            runs(i) = temp
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = doub(i + 1)
            doub(i + 1) = doub(i)
            doub(i) = temp
            temp = atbat(i + 1)
            atbat(i + 1) = atbat(i)
            atbat(i) = temp
            temp = hr(i + 1)
            hr(i + 1) = hr(i)
            hr(i) = temp
            temp = rbi(i + 1)
            rbi(i + 1) = rbi(i)
            rbi(i) = temp
            temp = stolenbase(i + 1)
            stolenbase(i + 1) = stolenbase(i)
            stolenbase(i) = temp
            temp = hits(i + 1)
            hits(i + 1) = hits(i)
            hits(i) = temp
        End If
    Next i
Next pass
For i = 1 To 19
    pbxResults.Print strName(i); Tab(25); trip(i)
Next i
Close #1


End Sub
Private Sub Form_Load()
    strpath = "N:\CS130\handin\Casadder\"
    Open strpath & "TwinsBatting.txt" For Input As #1
    For i = 1 To 19
    Input #1, strName(i), games(i), atbat(i), runs(i), hits(i), doub(i), trip(i), hr(i), rbi(i), stolenbase(i), avg(i)
    Next i
    Close #1
End Sub

