VERSION 5.00
Begin VB.Form frmRoseCompared 
   BackColor       =   &H8000000E&
   Caption         =   "Rose Compared Side-by-Side"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   1335
   ClientWidth     =   15240
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   15240
   Begin VB.PictureBox picResultsVersus 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11595
      TabIndex        =   4
      Top             =   5520
      Width           =   11655
   End
   Begin VB.TextBox txtSelectPlayer 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Height          =   735
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdReturnMenu 
      BackColor       =   &H000000FF&
      DisabledPicture =   "frmRoseCompared.frx":0000
      Height          =   1215
      Left            =   1320
      Picture         =   "frmRoseCompared.frx":7DDC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdCompareRose 
      BackColor       =   &H000000C0&
      Height          =   1935
      Left            =   720
      Picture         =   "frmRoseCompared.frx":F721
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   3015
   End
   Begin VB.PictureBox picResultsCompare 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000E&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   5520
      Width           =   11655
   End
   Begin VB.Label lblEnterNumber 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Enter the Number of the Hall of Famer You want to Compare to Pete Rose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Image picRoseBack 
      Height          =   8595
      Left            =   11280
      Picture         =   "frmRoseCompared.frx":1A361
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5250
   End
   Begin VB.Image picHOFnumbered 
      Height          =   5655
      Left            =   4920
      Picture         =   "frmRoseCompared.frx":2DD26
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6180
   End
End
Attribute VB_Name = "frmRoseCompared"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'declares variables to be used in a search of players for use in comparison
    Dim GP(1 To 30), AB(1 To 30), R(1 To 30), H(1 To 30), DBL(1 To 30), TRI(1 To 30), HR(1 To 30), RBI(1 To 30), BB(1 To 30), SO(1 To 30), SB(1 To 30), CS(1 To 30), BA(1 To 30) As Double
    Dim HOFname(1 To 30) As String
    Dim HOFyear(1 To 30), Seasons(1 To 30) As Double
    Dim CTR As Integer
Private Sub cmdCompareRose_Click()
    'searches players within the file \HoF_stats_withPete_numbered.txt for display and comparison
    picResultsCompare.Cls
    picResultsVersus.Cls
    Dim HOFnumber(1 To 30) As Integer
    Dim HOFselected(1 To 30) As Integer
    Dim CTR As Integer
    CTR = 0
    Open App.Path & "\HoF_stats_withPete_numbered.txt" For Input As #1
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, HOFnumber(CTR), HOFname(CTR), HOFyear(CTR), Seasons(CTR), GP(CTR), AB(CTR), R(CTR), H(CTR), DBL(CTR), TRI(CTR), HR(CTR), RBI(CTR), BB(CTR), SO(CTR), SB(CTR), CS(CTR), BA(CTR)
    Loop
    Close #1
    Dim searchName As Double
    searchName = Val(txtSelectPlayer.Text)
        Select Case searchName
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24
            Case searchName > 1
                MsgBox "Sorry, the number you entered is invalid. Please enter a number listed (1 - 24).", , "ERROR"
            Case searchName < 24
                MsgBox "Sorry, the number you entered is invalid. Please enter a number listed (1 - 24).", , "ERROR"
            Case Else
                MsgBox "Sorry, the number you entered is invalid. Please enter a number listed (1 - 24).", , "ERROR"
        End Select
    Dim count, counter As Integer
    Dim Found As Boolean
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFnumber(count) = searchName Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsVersus.Print " Pete Rose vs. "; HOFname(count)
            picResultsCompare.Print Chr(10); Chr(10); Chr(10); Chr(10); Chr(10); Tab(2); HOFname(CTR); Chr(10); Tab(2); "("; HOFyear(CTR); ", "; Seasons(CTR); "seasons )"; Tab(30); GP(CTR); Tab(40); AB(CTR); Tab(50); R(CTR); Tab(60); H(CTR); Tab(70); DBL(CTR); Tab(80); TRI(CTR); Tab(90); HR(CTR); Tab(100); RBI(CTR); Tab(110); BB(CTR); Tab(120); SO(CTR); Tab(130); SB(CTR); Tab(140); CS(CTR); Tab(150); Right(FormatNumber(BA(CTR), 3), 4);
            picResultsCompare.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
            picResultsCompare.Print Chr(10); "  Player"; Tab(2); "(Year Inducted,"; Tab(31); "G"; Tab(41); "AB"; Tab(51); "R"; Tab(61); "H"; Tab(71); "2B"; Tab(80); "3B"; Tab(91); "HR"; Tab(102); "RBI"; Tab(111); "BB"; Tab(121); "SO"; Tab(131); "SB"; Tab(141); "CS"; Tab(151); "BA"; Tab(2); "Seasons Played)";
            picResultsCompare.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
            picResultsCompare.Print Chr(10); Tab(2); HOFname(count); Chr(10); Tab(2); "("; HOFyear(count); ", "; Seasons(count); "seasons )"; Tab(30); GP(count); Tab(40); AB(count); Tab(50); R(count); Tab(60); H(count); Tab(70); DBL(count); Tab(80); TRI(count); Tab(90); HR(count); Tab(100); RBI(count); Tab(110); BB(count); Tab(120); SO(count); Tab(130); SB(count); Tab(140); CS(count); Tab(150); Right(FormatNumber(BA(count), 3), 4);
      
        
        End If
End Sub

Private Sub cmdReturnMenu_Click()
    'returns user to previous screen (CareerStats) and hides form RoseCompared from visibility
    frmRoseCompared.Hide
    frmCareerStats.Show

End Sub

