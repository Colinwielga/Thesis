VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF0000&
   Caption         =   "2003 Minnesota Twins Pitching Statistics by Charlie Sadder"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form2"
   ScaleHeight     =   7230
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMeetP 
      Caption         =   "Individual Pitcher Stats"
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdERA 
      Caption         =   "ERA"
      Height          =   495
      Left            =   9840
      TabIndex        =   13
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdSO9 
      Caption         =   "K/9"
      Height          =   495
      Left            =   9000
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmpSO 
      Caption         =   "SO"
      Height          =   495
      Left            =   8160
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdIP 
      Caption         =   "IP"
      Height          =   495
      Left            =   7320
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdSaves 
      Caption         =   "SV"
      Height          =   495
      Left            =   6480
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdLosses 
      Caption         =   "L"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdWins 
      Caption         =   "W"
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdGames 
      Caption         =   "G"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdLoadP 
      Caption         =   "All Pitching Statistics"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H000000FF&
      Height          =   5775
      Left            =   2280
      ScaleHeight     =   5715
      ScaleWidth      =   7875
      TabIndex        =   3
      Top             =   1440
      Width           =   7935
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSwitchB 
      Caption         =   "Switch to Batting Statistics"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Earned run average"
      Height          =   615
      Left            =   9840
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Strikouts per 9 Innings"
      Height          =   615
      Left            =   9000
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Strikeouts"
      Height          =   255
      Left            =   8160
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Innings Pitched"
      Height          =   495
      Left            =   7320
      TabIndex        =   18
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Saves"
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Losses"
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Winis"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Games"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select Category for Pitching Leaders:"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1890
      Left            =   0
      Picture         =   "Form2.frx":0000
      Top             =   3000
      Width           =   1890
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MinnesotaTwinsVB (Charlie Sadder's VB project)
'Form Name : 2003 Minnesota Twins Pitching Statistics (form2.frm)
'Author: Charlie Sadder
'Date Written: October 28, 2003
'Purpose of Form: To see the Twins' leaders of statistics in the 2003 season.
                  'There are different categories for pitching
                  'for which you can see team leaders.

'Option Explicit is a command to force
'the user to declare all variables before the program will run.


Option Explicit
Dim i As Integer
Dim strName(1 To 16) As String
Dim games(1 To 16) As Integer
Dim wins(1 To 16) As Integer
Dim losses(1 To 16) As Integer
Dim saves(1 To 16) As Integer
Dim innpitch(1 To 16) As Integer
Dim strikeouts(1 To 16) As Integer
Dim ksper9(1 To 16) As Single
Dim era(1 To 16) As Single
Private Sub cmdClear_Click()
    'clear anything in pbxResults
    pbxResults.Cls
End Sub

Private Sub cmdERA_Click()
'order the pitchers from lowest to highest ERA.
Dim N As Integer
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 16
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "ERA"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If era(i) > era(i + 1) Then
            tempor = era(i + 1)
            era(i + 1) = era(i)
            era(i) = tempor
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = wins(i + 1)
            wins(i + 1) = wins(i)
            wins(i) = temp
            temp = losses(i + 1)
            losses(i + 1) = losses(i)
            losses(i) = temp
            temp = saves(i + 1)
            saves(i + 1) = saves(i)
            saves(i) = temp
            temp = innpitch(i + 1)
            innpitch(i + 1) = innpitch(i)
            innpitch(i) = temp
            temp = strikeouts(i + 1)
            strikeouts(i + 1) = strikeouts(i)
            strikeouts(i) = temp
            tempor = ksper9(i + 1)
            ksper9(i + 1) = ksper9(i)
            ksper9(i) = tempor
        End If
    Next i
Next pass
For i = 1 To 16
    pbxResults.Print strName(i); Tab(25); era(i)
Next i
Close #2

End Sub

Private Sub cmdGames_Click()
'Order the pitchers from most to least amount of games played.
Dim N As Integer
Dim temp As Single
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 16
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
            tempor = era(i + 1)
            era(i + 1) = era(i)
            era(i) = tempor
            temp = wins(i + 1)
            wins(i + 1) = wins(i)
            wins(i) = temp
            temp = losses(i + 1)
            losses(i + 1) = losses(i)
            losses(i) = temp
            temp = saves(i + 1)
            saves(i + 1) = saves(i)
            saves(i) = temp
            temp = innpitch(i + 1)
            innpitch(i + 1) = innpitch(i)
            innpitch(i) = temp
            temp = strikeouts(i + 1)
            strikeouts(i + 1) = strikeouts(i)
            strikeouts(i) = temp
            tempor = ksper9(i + 1)
            ksper9(i + 1) = ksper9(i)
            ksper9(i) = tempor
        End If
    Next i
Next pass
For i = 1 To 16
    pbxResults.Print strName(i); Tab(25); games(i)
Next i
Close #2
End Sub

Private Sub cmdIP_Click()
'Orders the pitchers from most to least innings pitched.
Dim N As Integer
Dim temp As Integer
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 16
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "IP"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If innpitch(i) < innpitch(i + 1) Then
            temp = innpitch(i + 1)
            innpitch(i + 1) = innpitch(i)
            innpitch(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = wins(i + 1)
            wins(i + 1) = wins(i)
            wins(i) = temp
            temp = losses(i + 1)
            losses(i + 1) = losses(i)
            losses(i) = temp
            temp = saves(i + 1)
            saves(i + 1) = saves(i)
            saves(i) = temp
            tempor = era(i + 1)
            era(i + 1) = era(i)
            era(i) = tempor
            temp = strikeouts(i + 1)
            strikeouts(i + 1) = strikeouts(i)
            strikeouts(i) = temp
            tempor = ksper9(i + 1)
            ksper9(i + 1) = ksper9(i)
            ksper9(i) = tempor
        End If
    Next i
Next pass
For i = 1 To 16
    pbxResults.Print strName(i); Tab(25); games(i)
Next i
Close #2

End Sub

Private Sub cmdLoadP_Click()
'Loads the pitchers and their statistics to the program.
    pbxResults.Cls
    pbxResults.Print " Player", Tab(15); "G"; Tab(25); "W"; Tab(35); "L"; Tab(45); "SV"; Tab(55); "IP"; Tab(65); "SO"; Tab(75); "K/9"; Tab(85); "ERA"; Tab(95)
    For i = 1 To 16
        
        pbxResults.Print strName(i), Tab(15); games(i); Tab(25); wins(i); Tab(35); losses(i); Tab(45); saves(i); Tab(55); innpitch(i); Tab(65); strikeouts(i); Tab(75); ksper9(i); Tab(85); era(i); Tab(95)
    Next i
    Close #2
End Sub
Private Sub cmdLosses_Click()
'Orders the pitchers from least to most losses.
 Dim N As Integer
Dim temp As Single
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 16
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "L"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If losses(i) > losses(i + 1) Then
            temp = losses(i + 1)
            losses(i + 1) = losses(i)
            losses(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = wins(i + 1)
            wins(i + 1) = wins(i)
            wins(i) = temp
            temp = innpitch(i + 1)
            innpitch(i + 1) = innpitch(i)
            innpitch(i) = temp
            temp = saves(i + 1)
            saves(i + 1) = saves(i)
            saves(i) = temp
            tempor = era(i + 1)
            era(i + 1) = era(i)
            era(i) = tempor
            temp = strikeouts(i + 1)
            strikeouts(i + 1) = strikeouts(i)
            strikeouts(i) = temp
            tempor = ksper9(i + 1)
            ksper9(i + 1) = ksper9(i)
            ksper9(i) = tempor
        End If
    Next i
Next pass
For i = 1 To 16
    pbxResults.Print strName(i); Tab(25); losses(i)
Next i
Close #2

End Sub

Private Sub cmdMeetP_Click()
Dim Found As Boolean
Dim iPlayer As String
Dim N As Integer
N = 16
i = 0
'has user enter a name of a pitcher to look up in the txt
iPlayer = InputBox("Enter the name of the player you wish to find")
Found = False
'Clear whatever may be in pbxResults for repeated use.
pbxResults.Cls
'Prints header on top of the picture box
pbxResults.Print " Player", Tab(15); "G"; Tab(25); "W"; Tab(35); "L"; Tab(45); "SV"; Tab(55); "IP"; Tab(65); "SO"; Tab(75); "K/9"; Tab(85); "ERA"; Tab(95)
Do While i <= N - 1 And Found = False
    'counts the number of times you go through your program
    i = i + 1
    If iPlayer = strName(i) Then
    Found = True
    End If
    Loop
If Found = True Then
        'prints out the name of the pitcher you entered and their stats
        pbxResults.Print strName(i), Tab(15); games(i); Tab(25); wins(i); Tab(35); losses(i); Tab(45); saves(i); Tab(55); innpitch(i); Tab(65); strikeouts(i); Tab(75); ksper9(i); Tab(85); era(i); Tab(95)
    Else
    'gives a pop up message and tells them there is no Pitcher with that name
        MsgBox ("Player not found")
End If
End Sub

Private Sub cmdQuit_Click()
'Quits the program.
    End
End Sub

Private Sub cmdSaves_Click()
'Orders the pitchers from most to least saves.
Dim N As Integer
Dim temp As Single
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 16
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "SV"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If saves(i) < saves(i + 1) Then
            temp = saves(i + 1)
            saves(i + 1) = saves(i)
            saves(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = wins(i + 1)
            wins(i + 1) = wins(i)
            wins(i) = temp
            temp = losses(i + 1)
            losses(i + 1) = losses(i)
            losses(i) = temp
            temp = innpitch(i + 1)
            innpitch(i + 1) = innpitch(i)
            innpitch(i) = temp
            tempor = era(i + 1)
            era(i + 1) = era(i)
            era(i) = tempor
            temp = strikeouts(i + 1)
            strikeouts(i + 1) = strikeouts(i)
            strikeouts(i) = temp
            tempor = ksper9(i + 1)
            ksper9(i + 1) = ksper9(i)
            ksper9(i) = tempor
        End If
    Next i
Next pass
For i = 1 To 16
    pbxResults.Print strName(i); Tab(25); saves(i)
Next i
Close #2
End Sub

Private Sub cmdSO9_Click()
'Orders the pitchers from most to least strikeouts per 9 innings
Dim N As Integer
Dim temp As Single
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 16
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "K/9"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If ksper9(i) < ksper9(i + 1) Then
            tempor = ksper9(i + 1)
            ksper9(i + 1) = ksper9(i)
            ksper9(i) = tempor
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = wins(i + 1)
            wins(i + 1) = wins(i)
            wins(i) = temp
            temp = losses(i + 1)
            losses(i + 1) = losses(i)
            losses(i) = temp
            temp = saves(i + 1)
            saves(i + 1) = saves(i)
            saves(i) = temp
            tempor = era(i + 1)
            era(i + 1) = era(i)
            era(i) = tempor
            temp = strikeouts(i + 1)
            strikeouts(i + 1) = strikeouts(i)
            strikeouts(i) = temp
            temp = innpitch(i + 1)
            innpitch(i + 1) = innpitch(i)
            innpitch(i) = temp
        End If
    Next i
Next pass
For i = 1 To 16
    pbxResults.Print strName(i); Tab(25); ksper9(i)
Next i
Close #2

End Sub

Private Sub cmdSwitchB_Click()
'switches from form 2 to form 3.
    Form2.Hide
    Form3.Show
End Sub
Private Sub cmdWins_Click()
'Orders the pitchers from most to least wins.
Dim N As Integer
Dim temp As Single
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 16
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "W"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If wins(i) < wins(i + 1) Then
            temp = wins(i + 1)
            wins(i + 1) = wins(i)
            wins(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = innpitch(i + 1)
            innpitch(i + 1) = innpitch(i)
            innpitch(i) = temp
            temp = losses(i + 1)
            losses(i + 1) = losses(i)
            losses(i) = temp
            temp = saves(i + 1)
            saves(i + 1) = saves(i)
            saves(i) = temp
            tempor = era(i + 1)
            era(i + 1) = era(i)
            era(i) = tempor
            temp = strikeouts(i + 1)
            strikeouts(i + 1) = strikeouts(i)
            strikeouts(i) = temp
            tempor = ksper9(i + 1)
            ksper9(i + 1) = ksper9(i)
            ksper9(i) = tempor
        End If
    Next i
Next pass
For i = 1 To 16
    pbxResults.Print strName(i); Tab(25); wins(i)
Next i
Close #2
End Sub

Private Sub cmpSO_Click()
'Orders the pitchers from most to least total strikeouts.
Dim N As Integer
Dim temp As Single
Dim pass As Integer
Dim tempquantity As String
Dim tempor As Double
N = 16
pbxResults.Cls
 pbxResults.Print " Player", Tab(25); "SO"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If strikeouts(i) < strikeouts(i + 1) Then
            temp = strikeouts(i + 1)
            strikeouts(i + 1) = strikeouts(i)
            strikeouts(i) = temp
            tempquantity = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = tempquantity
            temp = games(i + 1)
            games(i + 1) = games(i)
            games(i) = temp
            temp = wins(i + 1)
            wins(i + 1) = wins(i)
            wins(i) = temp
            temp = losses(i + 1)
            losses(i + 1) = losses(i)
            losses(i) = temp
            temp = saves(i + 1)
            saves(i + 1) = saves(i)
            saves(i) = temp
            tempor = era(i + 1)
            era(i + 1) = era(i)
            era(i) = tempor
            temp = innpitch(i + 1)
            innpitch(i + 1) = innpitch(i)
            innpitch(i) = temp
            tempor = ksper9(i + 1)
            ksper9(i + 1) = ksper9(i)
            ksper9(i) = tempor
        End If
    Next i
Next pass
For i = 1 To 16
    pbxResults.Print strName(i); Tab(25); strikeouts(i)
Next i
Close #2

End Sub

Private Sub Form_Load()
    strpath = "N:\CS130\handin\Casadder\"
    Open strpath & "TwinsPitching.txt" For Input As #2
    For i = 1 To 16
        Input #2, strName(i), games(i), wins(i), losses(i), saves(i), innpitch(i), strikeouts(i), ksper9(i), era(i)
    Next i
    Close #2
End Sub
