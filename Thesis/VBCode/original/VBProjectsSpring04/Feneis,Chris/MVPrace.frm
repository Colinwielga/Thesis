VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Most Valuable Player Race By: Chris Feneis"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNextyear 
      BackColor       =   &H00008080&
      Caption         =   "Ahhh...Better luck next year"
      Height          =   975
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00008080&
      Caption         =   "Whose name has more than 10 Characters?"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   3495
   End
   Begin VB.CommandButton cmdBA 
      BackColor       =   &H000000FF&
      Caption         =   "Whose Batting Average is this?"
      Height          =   735
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H0000FF00&
      Caption         =   "Put them in order according to the MVP race"
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton cmdMVP 
      BackColor       =   &H00008080&
      Caption         =   "And our MVP is..."
      Height          =   1215
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlace 
      BackColor       =   &H00FF0000&
      Caption         =   "Display How many MVP points they have"
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdRankings 
      BackColor       =   &H00008080&
      Caption         =   "Read and Display Rankings"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00008080&
      Caption         =   "Quit"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H000000FF&
      Caption         =   "Read and Display Stats"
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3675
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label Label3 
      Caption         =   "Start Here   ==>"
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Open Read and Display Stats First"
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Open Read and Display Stats First"
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Most Valuable Player Race (MVPrace.vbp)
'Most Valuable Player Race (MVPrace.frm)
'By: Chris Feneis
'March 11-15, 2004
'I wrote the code the way that i did because this is how i learned how to write
'the code.  This was the easiest for me to understand, and for you to understand
'also.
'I chose to do this project because I am a Baseball fanatic.  I love everything
'about Baseball and I watch the stats closely, so I wanted to find out who was
'the MVP last year according to these 7 key stats.  I chose these 7 categories
'because i felt these were the most important to the team, not just the individual.
'This project was very fun to do because I was able to use something that I'm good
'at and applied it to something that I wasn't quite as good at, and I became better
'at the visual basics because of this project.  It was a fun time.
Option Explicit
Dim Names(1 To 8) As String, OBP(1 To 8) As String, Path As String
Dim R(1 To 8) As String, SLG(1 To 8) As String, HR(1 To 8) As String
Dim RBI(1 To 8) As String, SB(1 To 8) As String, AVE(1 To 8) As String
Dim J As Single, MVPpoints(1 To 8) As String, BattingAverage As Single
Private Sub cmdBA_Click()
Close #1
'open the message box( one of the 10)
MsgBox "You must use a period followed by 3 decimal places.", , "Warning"
Dim J As Integer, Found As Boolean, BattingAve As Integer
BattingAverage = InputBox("Enter the Batting Average", BattingAve)
Found = False
J = 0
    'making the loop (& making sure it doesn't go out of range)
    Do While (Not Found) And (J < 8)
    J = J + 1
    If BattingAverage = AVE(J) Then
    Found = True
End If
Loop
    If Found Then
    picResults.Print Names(J); " hit"; AVE(J); " this year."
    Else
    picResults.Print "Sorry, that was not one of the MVP candidates Batting Averages!"
    End If
End Sub

Private Sub cmdMVP_Click()
    picResults.Cls
    picResults.Picture = LoadPicture(Path & "pujols.jpg", vbLPLarge, vbLPColor)

    picResults.Print Tab(15); "A.Pujols is the 2003 MLB MVP!!!"
cmdMVP.Enabled = False
End Sub

Private Sub cmdName_Click()
    picResults.Cls
Dim namelength As String, X As Integer
namelength = Len(Names(1))
    X = Len("Matsui")
    X = 6
    'keeping it within the range
    For J = 1 To 8
    'searching through the data & looking for characters of 10 +
    If Len(Names(J)) > 10 Then
    picResults.Print Names(J)
    End If
Next J
cmdName.Enabled = False
End Sub

Private Sub cmdNextyear_Click()
    picResults.Picture = LoadPicture(Path & "MVPpix.jpg.bmp", vbLPLarge, vbLPColor)
    'making space for the text to come up clearly(could've just used tab, but didn't)
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print "Sorry, Mr.Sheffield, Mr. Rodriguez, Mr. Helton, Mr. Delgado, Mr. Soriano, Mr. Bonds, and"
    picResults.Print "Mr. Ramirez...You have to try harder next year"
cmdNextyear.Enabled = False
End Sub

Private Sub cmdOrder_Click()
Close #1
    picResults.Print
Dim PASS As Integer, COMP As Integer, tempNames As String, tempMVPpoints As String
Dim Ctr As Integer, J As Integer, Names(1 To 8) As String, MVPpoints(1 To 8) As String
Open Path & "Points.txt" For Input As #1
Ctr = 8
    For J = 1 To Ctr
Input #1, Names(J), MVPpoints(J)
Next J

For PASS = 1 To Ctr - 1
    For COMP = 1 To Ctr - PASS
    If MVPpoints(COMP) < MVPpoints(COMP + 1) Then
        
        'switches the points around
        tempMVPpoints = MVPpoints(COMP)
        MVPpoints(COMP) = MVPpoints(COMP + 1)
        MVPpoints(COMP + 1) = tempMVPpoints
        
        'switches the names around
        tempNames = Names(COMP)
        Names(COMP) = Names(COMP + 1)
        Names(COMP + 1) = tempNames
        
    End If
  Next COMP
 Next PASS
 
 For J = 1 To Ctr
    picResults.Print Names(J), MVPpoints(J)
 Next J
cmdOrder.Enabled = False
End Sub
Private Sub cmdPlace_Click()
picResults.Cls
Close #1
Open Path & "Points.txt" For Input As #1
    'keeping the range so no errors come up
    For J = 1 To 8
Input #1, Names(J), MVPpoints(J)
 Next J
J = 0
    'setting a header so it is easier to understand
    picResults.Print "Names"; Tab(15); "MVPpoints"
    'Just reading the data and eventually displaying the data
    Do While J < 8
    J = J + 1
    picResults.Print Names(J); Tab(15); MVPpoints(J)
Loop
cmdPlace.Enabled = False
End Sub

Private Sub cmdRankings_Click()
Close #1
'Opening our data so we can access it in the program
Open Path & "Rankings.txt" For Input As #1
    'range is not going to be an error
    For J = 1 To 8
Input #1, Names(J), OBP(J), R(J), SLG(J), HR(J), RBI(J), SB(J), AVE(J)
 Next J
J = 0
    picResults.Print
    Do While J < 8
    J = J + 1
    picResults.Print Names(J); Tab(15); OBP(J); Tab(21); R(J); Tab(26); SLG(J); Tab(32); HR(J); Tab(37); RBI(J); Tab(42); SB(J); Tab(47); AVE(J)
Loop
cmdRankings.Enabled = False
End Sub

Private Sub cmdRead_Click()
    'clearing out the screen each time to make room for the next
    picResults.Cls
Close #1
    'header to make things make more sense
    picResults.Print "Names"; Tab(15); "OBP"; Tab(21); "R"; Tab(26); "SLG"; Tab(32); "HR"; Tab(37); "RBI"; Tab(42); "SB"; Tab(47); "AVE"
Open Path & "Stats.txt" For Input As #1
    For J = 1 To 8
Input #1, Names(J), OBP(J), R(J), SLG(J), HR(J), RBI(J), SB(J), AVE(J)
 Next J
J = 0
    'reading the data and printing it out for our viewers
    Do While J < 8
    J = J + 1
    picResults.Print Names(J); Tab(15); OBP(J); Tab(21); R(J); Tab(26); SLG(J); Tab(32); HR(J); Tab(37); RBI(J); Tab(42); SB(J); Tab(47); AVE(J)
Loop
End Sub
Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
'setting it to path so my project can actually work for you
Path = "N:\CS130\handin\Feneis, Chris\"
End Sub
