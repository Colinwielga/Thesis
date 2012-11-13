VERSION 5.00
Begin VB.Form frmStatistics 
   BackColor       =   &H00800000&
   Caption         =   "Statistics"
   ClientHeight    =   7830
   ClientLeft      =   2715
   ClientTop       =   1665
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   Picture         =   "frmCareerStats0005.frx":0000
   ScaleHeight     =   7830
   ScaleWidth      =   10005
   Visible         =   0   'False
   Begin VB.CommandButton cmdMain4 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   6
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H8000000D&
      Caption         =   "Calculate"
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000013&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   9675
      TabIndex        =   3
      Top             =   2520
      Width           =   9735
   End
   Begin VB.CommandButton cmdCareerTotal 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Career Total   "
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7080
      Picture         =   "frmCareerStats0005.frx":EA64E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdStats0005 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Career Stats 2000-2005"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4200
      Picture         =   "frmCareerStats0005.frx":EE394
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdStats9500 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Career Stats 1995-2000"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1080
      Picture         =   "frmCareerStats0005.frx":F20DA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblCalc 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Use this to Calculate Averages from more than one year "
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3360
      TabIndex        =   5
      Top             =   6480
      Width           =   1815
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ProjectKG
'frmStatistics
'Jon Jerabek
'10-25-05 & 10-26-05
'Objective-Allows user to view stats and calculate an average

Private Sub cmdCalculate_Click()
Dim x As Double
Dim y As Double
Dim Sum As Double
Dim Average As Double
Sum = 0
Average = 0
x = InputBox("Enter How Many Stats You Are Averaging", "Averages")  'User inputs # of stats averaging
For y = 1 To x
    Sum = InputBox("Enter Stat", "Averages") + Sum  'User inputs stats
Next y
Average = Sum / x
MsgBox FormatNumber(Average, 3), , "Here Is Your Average"   'The numbers are added and divided by x


End Sub

Private Sub cmdCareerTotal_Click()  'Opens file containing stats and displays career average
picOutput.Cls
Open App.Path & "\CareerStat.txt" For Input As #1
For I = 1 To 1
    Input #1, Career(I), G(I), GS(I), MPG(I), FG(1), Three(I), FT(I), OReb(I), DReb(I), RPG(I), APG(I), SPG(I), BPG(I), TurnO(I), PF(I), PPG(I)
Next I
Close #1

For I = 1 To 1
    picOutput.Print " "; Tab(7); "G"; Tab(14); "GS"; Tab(21); "MPG"; Tab(28); "FG"; Tab(35); "Three"; Tab(42); "FT"; Tab(49); "OReb"; Tab(56); "DReb"; Tab(63); "RPG"; Tab(70); "APG"; Tab(77); "SPG"; Tab(84); "BPG"; Tab(91); "TurnO"; Tab(98); "PF"; Tab(105); "PPG"
    picOutput.Print " "; Tab(7); G(I); Tab(14); GS(I); Tab(21); MPG(I); Tab(28); FG(1); Tab(35); Three(I); Tab(42); FT(I); Tab(49); OReb(I); Tab(56); DReb(I); Tab(63); RPG(I); Tab(70); APG(I); Tab(77); SPG(I); Tab(84); BPG(I); Tab(91); TurnO(I); Tab(98); PF(I); Tab(105); PPG(I)
Next I
End Sub

Private Sub cmdMain4_Click()
frmHome.Show
frmStatistics.Hide
End Sub

Private Sub cmdStats0005_Click()    'Opens file containing stats and displays last 5 years
picOutput.Cls
Open App.Path & "\Stats.txt" For Input As #3
For I = 1 To 10
    Input #3, Yr(I), G(I), GS(I), MPG(I), FG(1), Three(I), FT(I), OReb(I), DReb(I), RPG(I), APG(I), SPG(I), BPG(I), TurnO(I), PF(I), PPG(I)
Next I
Close #3
picOutput.Print "Yr"; Tab(7); "G"; Tab(14); "GS"; Tab(21); "MPG"; Tab(28); "FG"; Tab(35); "Three"; Tab(42); "FT"; Tab(49); "OReb"; Tab(56); "DReb"; Tab(63); "RPG"; Tab(70); "APG"; Tab(77); "SPG"; Tab(84); "BPG"; Tab(91); "TurnO"; Tab(98); "PF"; Tab(105); "PPG"
For I = 6 To 10
    picOutput.Print Yr(I); Tab(7); G(I); Tab(14); GS(I); Tab(21); MPG(I); Tab(28); FG(1); Tab(35); Three(I); Tab(42); FT(I); Tab(49); OReb(I); Tab(56); DReb(I); Tab(63); RPG(I); Tab(70); APG(I); Tab(77); SPG(I); Tab(84); BPG(I); Tab(91); TurnO(I); Tab(98); PF(I); Tab(105); PPG(I)
Next I

End Sub

Private Sub cmdStats9500_Click()    'Opens file containing stats and displays first 5 years
picOutput.Cls
Open App.Path & "\Stats.txt" For Input As #2
For I = 1 To 10
    Input #2, Yr(I), G(I), GS(I), MPG(I), FG(1), Three(I), FT(I), OReb(I), DReb(I), RPG(I), APG(I), SPG(I), BPG(I), TurnO(I), PF(I), PPG(I)
Next I
Close #2
picOutput.Print "Yr"; Tab(7); "G"; Tab(14); "GS"; Tab(21); "MPG"; Tab(28); "FG"; Tab(35); "Three"; Tab(42); "FT"; Tab(49); "OReb"; Tab(56); "DReb"; Tab(63); "RPG"; Tab(70); "APG"; Tab(77); "SPG"; Tab(84); "BPG"; Tab(91); "TurnO"; Tab(98); "PF"; Tab(105); "PPG"
For I = 1 To 5
    picOutput.Print Yr(I); Tab(7); G(I); Tab(14); GS(I); Tab(21); MPG(I); Tab(28); FG(1); Tab(35); Three(I); Tab(42); FT(I); Tab(49); OReb(I); Tab(56); DReb(I); Tab(63); RPG(I); Tab(70); APG(I); Tab(77); SPG(I); Tab(84); BPG(I); Tab(91); TurnO(I); Tab(98); PF(I); Tab(105); PPG(I)
Next I
End Sub
