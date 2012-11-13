VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Most Valuable Player Race By: Chris Feneis"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
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
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "Open Read and Display Data First"
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
Dim Names(1 To 8) As String, OBP(1 To 8) As String, Path As String
Dim R(1 To 8) As String, SLG(1 To 8) As String, HR(1 To 8) As String
Dim RBI(1 To 8) As String, SB(1 To 8) As String, AVE(1 To 8) As String
Dim J As Single, MVPpoints(1 To 8) As String, BattingAverage As Single
Private Sub cmdBA_Click()
Close #1
MsgBox "You must open Read and Display Stats first, so type in a number, then try again later", , "Warning"
Dim J As Integer, Found As Boolean
BattingAverage = InputBox("Enter the Batting Average", BattingAve)
Found = False
J = 0
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
Dim namelength As String
namelength = Len(Names(1))
    X = Len("Matsui")
    X = 6
    For J = 1 To 8
    If Len(Names(J)) > 10 Then
    picResults.Print Names(J)
    End If
Next J
cmdName.Enabled = False
End Sub

Private Sub cmdNextyear_Click()
picResults.Picture = LoadPicture(Path & "MVPpix.jpg.bmp", vbLPLarge, vbLPColor)
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
Dim PASS As Integer, COMP As Integer, tempNames As Single
Dim Ctr As Integer, listtotal As Integer, J As Integer
Open Path & "Points.txt" For Input As #1
For J = 1 To Ctr
    Input #1, Names(J), MVPpoints(J)
Next J

For PASS = 1 To listtotal - 1
    For J = 1 To Ctr - PASS
    If MVPpoints(J) < MVPpoints(J + 1) Then
        tempNames = Names(J)
        Names(J) = Names(J + 1)
        Names(J + 1) = tempNames
    End If
  Next J
 Next PASS
 
    For J = 1 To Ctr
        picResults.Print Names(J)
    Next J
cmdOrder.Enabled = False
 End Sub
Private Sub cmdPlace_Click()
picResults.Cls
Close #1
Open Path & "Points.txt" For Input As #1
For J = 1 To 8
Input #1, Names(J), MVPpoints(J)
Next J
J = 0
    picResults.Print "Names"; Tab(15); "MVPpoints"
    Do While J < 8
    J = J + 1
    picResults.Print Names(J); Tab(15); MVPpoints(J)
Loop
cmdPlace.Enabled = False
End Sub

Private Sub cmdRankings_Click()
Close #1
Open Path & "Rankings.txt" For Input As #1
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
picResults.Cls
Close #1
picResults.Print "Names"; Tab(15); "OBP"; Tab(21); "R"; Tab(26); "SLG"; Tab(32); "HR"; Tab(37); "RBI"; Tab(42); "SB"; Tab(47); "AVE"
Open Path & "Stats.txt" For Input As #1
For J = 1 To 8
Input #1, Names(J), OBP(J), R(J), SLG(J), HR(J), RBI(J), SB(J), AVE(J)
Next J
J = 0
    Do While J < 8
    J = J + 1
    picResults.Print Names(J); Tab(15); OBP(J); Tab(21); R(J); Tab(26); SLG(J); Tab(32); HR(J); Tab(37); RBI(J); Tab(42); SB(J); Tab(47); AVE(J)
Loop
End Sub
Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
Path = "M:\labs\Feneis, Chris\"
End Sub
