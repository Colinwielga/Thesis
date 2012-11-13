VERSION 5.00
Begin VB.Form frmCounties 
   BackColor       =   &H80000009&
   Caption         =   "Republic of Srpska, election 2006"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF8080&
      Caption         =   "Back to main page"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9360
      Width           =   2055
   End
   Begin VB.TextBox txtEnterName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   20
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdMN6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Arrange by number of chairs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdMN2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Arrange by number of chairs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdPN3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Arrange  by Party name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdMN3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Arrange by number of chairs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPN4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Arrange  by Party name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdMN4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Arrange by number of chairs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPN5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Arrange  by Party name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdMN5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Arrange by number of chairs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdPN6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Arrange  by Party name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdPN2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Arrange  by Party name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdMN1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Arrange by number of chairs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdPN1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Arrange  by Party name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCoutny3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Research County 3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdCounty5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Research County 5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCounty4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Research County 4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdCounty6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Research County 6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCouty2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Research County 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdCounty1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Research County 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdFinal 
      BackColor       =   &H00FF8080&
      Caption         =   "Final Results"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   5055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10095
      Left            =   6600
      ScaleHeight     =   10035
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   840
      Width           =   8175
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "please use capital letters"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblEnterName 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Please enter name of wished Party in Text Box below "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   21
      Top             =   1680
      Width           =   3015
   End
End
Attribute VB_Name = "frmCounties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This code defineds Global Values
Dim Sum As Double, Percent(1 To 30) As Single, I As Integer, total(1 To 30) As Single
Dim Pass As Integer, Pos As Integer, Temp As String, Temp1 As Single, Found As Boolean, InputName As String

Private Sub cmdBack_Click()

'Cod Option that connect two forms
'In this case this form and Main Page
frmMainPage.Show
frmCounties.Hide

End Sub

Private Sub cmdCounty1_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 1"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County1.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr1 = 0
Sum = 0
Do While Not EOF(1)
Ctr1 = Ctr1 + 1
Input #1, PartyName(Ctr1), Votes(Ctr1)
Sum = Sum + Votes(Ctr1)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr1
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input

End Sub

Private Sub cmdCounty4_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 4"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County4.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr4 = 0
Sum = 0
Do While Not EOF(1)
Ctr4 = Ctr4 + 1
Input #1, PartyName(Ctr4), Votes(Ctr4)
Sum = Sum + Votes(Ctr4)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr3
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input


End Sub

Private Sub cmdCounty5_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 5"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County5.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr5 = 0
Sum = 0
Do While Not EOF(1)
Ctr5 = Ctr5 + 1
Input #1, PartyName(Ctr5), Votes(Ctr5)
Sum = Sum + Votes(Ctr5)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr5
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input

End Sub

Private Sub cmdCounty6_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 6"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County6.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr6 = 0
Sum = 0
Do While Not EOF(1)
Ctr6 = Ctr6 + 1
Input #1, PartyName(Ctr6), Votes(Ctr6)
Sum = Sum + Votes(Ctr6)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr6
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input

End Sub

Private Sub cmdCoutny3_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 3"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County3.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr3 = 0
Sum = 0
Do While Not EOF(1)
Ctr3 = Ctr3 + 1
Input #1, PartyName(Ctr3), Votes(Ctr3)
Sum = Sum + Votes(Ctr3)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr3
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input

End Sub

Private Sub cmdCouty2_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 2"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County2.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr2 = 0
Sum = 0
Do While Not EOF(1)
Ctr2 = Ctr2 + 1
Input #1, PartyName(Ctr2), Votes(Ctr2)
Sum = Sum + Votes(Ctr2)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr2
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input

End Sub

Private Sub cmdFinal_Click()

picResults.Cls

Dim Ctr7 As Integer, Ctr8 As Integer, Ctr9 As Integer, Ctr10 As Integer, Ctr11 As Integer, Ctr12 As Integer
Dim Ctr13 As Integer, Ctr14 As Integer, Ctr15 As Integer, Ctr16 As Integer, Ctr17 As Integer, Ctr18 As Integer
Dim Ctr19 As Integer, Ctr20 As Integer, Ctr21 As Integer, Ctr22 As Integer

Dim Sum7 As Double, Sum8 As Double, Sum9 As Double, Sum10 As Double
Dim Sum11 As Double, Sum12 As Double, Sum13 As Double, Sum14 As Double
Dim Sum15 As Double, Sum16 As Double, Sum17 As Double, Sum18 As Double
Dim Sum19 As Double, Sum20 As Double, Sum21 As Double, Sum22 As Double

Dim Percents7 As Single, Percents8 As Single, Percents9 As Single, Percents10 As Single, Percents11 As Single
Dim Percents12 As Single, Percents13 As Single, Percents14 As Single, Percents15 As Single, Percents16 As Single
Dim Percents17 As Single, Percents18 As Single, Percents19 As Single, Percents20 As Single, Percents21 As Single, Percents22 As Single


Open App.Path & "\SNSD.txt" For Input As #7
Open App.Path & "\SDS.txt" For Input As #8
Open App.Path & "\PDP-RS.txt" For Input As #9
Open App.Path & "\DNS.txt" For Input As #10
Open App.Path & "\SzBH.txt" For Input As #11
Open App.Path & "\SP.txt" For Input As #12
Open App.Path & "\SDA.txt" For Input As #13
Open App.Path & "\SRS-RS.txt" For Input As #14
Open App.Path & "\PS-RS.txt" For Input As #15
Open App.Path & "\SDP.txt" For Input As #16
Open App.Path & "\SRS-VS.txt" For Input As #17
Open App.Path & "\DEPOS.txt" For Input As #18
Open App.Path & "\NS-RZB.txt" For Input As #19
Open App.Path & "\NSS.txt" For Input As #20
Open App.Path & "\NHI-HDZ-HSP-HNZ.txt" For Input As #21
Open App.Path & "\DSS.txt" For Input As #22


picResults.Print Tab(3); "Final Results for Election 2006 in RS"
picResults.Print

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr7 = 0
Sum7 = 0
Do While Not EOF(7)
Ctr7 = Ctr7 + 1
Input #7, CountyName(Ctr7), PartyName(Ctr7), Votes(Ctr7)
Sum7 = Sum7 + Votes(Ctr7)
Loop

'This coman calculate percent of von votes
Percents7 = Sum7 / TotalCorectVotes

picResults.Print Tab(3); PartyName(Ctr7); " won "; Sum7; " votes. That means   "; FormatPercent(Percents7, 2)


'This array read information from specific input and gave
'us number of von votes for specific party
Ctr8 = 0
Sum8 = 0
Do While Not EOF(8)
Ctr8 = Ctr8 + 1
Input #8, CountyName(Ctr8), PartyName(Ctr8), Votes(Ctr8)
Sum8 = Sum8 + Votes(Ctr8)
Loop

'This coman calculate percent of von votes
Percents8 = Sum8 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr8); " won "; Sum8; " votes. That means   "; FormatPercent(Percents8, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr9 = 0
Sum9 = 0
Do While Not EOF(9)
Ctr9 = Ctr9 + 1
Input #9, CountyName(Ctr9), PartyName(Ctr9), Votes(Ctr9)
Sum9 = Sum9 + Votes(Ctr9)
Loop

'This coman calculate percent of von votes
Percents9 = Sum9 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr9); " won "; Sum9; " votes. That means   "; FormatPercent(Percents9, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr10 = 0
Sum10 = 0
Do While Not EOF(10)
Ctr10 = Ctr10 + 1
Input #10, CountyName(Ctr10), PartyName(Ctr10), Votes(Ctr10)
Sum10 = Sum10 + Votes(Ctr10)
Loop

'This coman calculate percent of von votes
Percents10 = Sum10 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr10); " won "; Sum10; " votes. That means   "; FormatPercent(Percents10, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr11 = 0
Sum11 = 0
Do While Not EOF(11)
Ctr11 = Ctr11 + 1
Input #11, CountyName(Ctr11), PartyName(Ctr11), Votes(Ctr11)
Sum11 = Sum11 + Votes(Ctr11)
Loop

'This coman calculate percent of von votes
Percents11 = Sum11 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr11); " won "; Sum11; " votes. That means   "; FormatPercent(Percents11, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr12 = 0
Sum12 = 0
Do While Not EOF(12)
Ctr12 = Ctr12 + 1
Input #12, CountyName(Ctr12), PartyName(Ctr12), Votes(Ctr12)
Sum12 = Sum12 + Votes(Ctr12)
Loop

'This coman calculate percent of von votes
Percents12 = Sum12 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr12); " won "; Sum12; " votes. That means   "; FormatPercent(Percents12, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr13 = 0
Sum13 = 0
Do While Not EOF(13)
Ctr13 = Ctr13 + 1
Input #13, CountyName(Ctr13), PartyName(Ctr13), Votes(Ctr13)
Sum13 = Sum13 + Votes(Ctr13)
Loop

'This coman calculate percent of von votes
Percents13 = Sum13 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr13); " won "; Sum13; " votes. That means   "; FormatPercent(Percents13, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr14 = 0
Sum14 = 0
Do While Not EOF(14)
Ctr14 = Ctr14 + 1
Input #14, CountyName(Ctr14), PartyName(Ctr14), Votes(Ctr14)
Sum14 = Sum14 + Votes(Ctr14)
Loop

'This coman calculate percent of von votes
Percents14 = Sum14 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr14); " won "; Sum14; " votes. That means   "; FormatPercent(Percents14, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr15 = 0
Sum15 = 0
Do While Not EOF(15)
Ctr15 = Ctr15 + 1
Input #15, CountyName(Ctr15), PartyName(Ctr15), Votes(Ctr15)
Sum15 = Sum15 + Votes(Ctr15)
Loop

'This coman calculate percent of von votes
Percents15 = Sum15 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr15); " won "; Sum15; " votes. That means   "; FormatPercent(Percents15, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr16 = 0
Sum16 = 0
Do While Not EOF(16)
Ctr16 = Ctr16 + 1
Input #16, CountyName(Ctr16), PartyName(Ctr16), Votes(Ctr16)
Sum16 = Sum16 + Votes(Ctr16)
Loop

'This coman calculate percent of von votes
Percents16 = Sum16 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr16); " won "; Sum16; " votes. That means   "; FormatPercent(Percents16, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr17 = 0
Sum17 = 0
Do While Not EOF(17)
Ctr17 = Ctr17 + 1
Input #17, CountyName(Ctr17), PartyName(Ctr17), Votes(Ctr17)
Sum17 = Sum17 + Votes(Ctr17)
Loop

'This coman calculate percent of von votes
Percents17 = Sum17 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr17); " won "; Sum17; " votes. That means   "; FormatPercent(Percents17, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr18 = 0
Sum18 = 0
Do While Not EOF(18)
Ctr18 = Ctr18 + 1
Input #18, CountyName(Ctr18), PartyName(Ctr18), Votes(Ctr18)
Sum18 = Sum18 + Votes(Ctr18)
Loop

'This coman calculate percent of von votes
Percents18 = Sum18 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr18); " won "; Sum18; " votes. That means   "; FormatPercent(Percents18, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr19 = 0
Sum19 = 0
Do While Not EOF(19)
Ctr19 = Ctr19 + 1
Input #19, CountyName(Ctr19), PartyName(Ctr19), Votes(Ctr19)
Sum19 = Sum19 + Votes(Ctr19)
Loop

'This coman calculate percent of von votes
Percents19 = Sum19 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr19); " won "; Sum19; " votes. That means   "; FormatPercent(Percents19, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr20 = 0
Sum20 = 0
Do While Not EOF(20)
Ctr20 = Ctr20 + 1
Input #20, CountyName(Ctr20), PartyName(Ctr20), Votes(Ctr20)
Sum20 = Sum20 + Votes(Ctr20)
Loop

'This coman calculate percent of von votes
Percents20 = Sum20 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr20); " won "; Sum20; " votes. That means   "; FormatPercent(Percents20, 2)

'This array read information from specific input and gave
'us number of von votes for specific party
Ctr21 = 0
Sum21 = 0
Do While Not EOF(21)
Ctr21 = Ctr21 + 1
Input #21, CountyName(Ctr21), PartyName(Ctr21), Votes(Ctr21)
Sum21 = Sum21 + Votes(Ctr21)
Loop

'This coman calculate percent of von votes
Percents21 = Sum21 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr21); " won "; Sum21; " votes. That means   "; FormatPercent(Percents21, 2)



'This array read information from specific input and gave
'us number of von votes for specific party
Ctr22 = 0
Sum22 = 0
Do While Not EOF(22)
Ctr22 = Ctr22 + 1
Input #22, CountyName(Ctr22), PartyName(Ctr22), Votes(Ctr22)
Sum22 = Sum22 + Votes(Ctr22)
Loop

'This coman calculate percent of von votes
Percents22 = Sum22 / TotalCorectVotes

picResults.Print
picResults.Print Tab(3); PartyName(Ctr22); " won "; Sum22; " votes. That means   "; FormatPercent(Percents22, 2)

Close #7
Close #8
Close #9
Close #10
Close #11
Close #12
Close #13
Close #14
Close #15
Close #16
Close #17
Close #18
Close #19
Close #20
Close #21
Close #22
End Sub

Private Sub cmdMN1_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 1"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County1.txt" For Input As #1

InputName = txtEnterName.Text

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr1 = 0
Sum = 0
Do While Not EOF(1)
Ctr1 = Ctr1 + 1
Input #1, PartyName(Ctr1), Votes(Ctr1)
Sum = Sum + Votes(Ctr1)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr1
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

'Option make empti space in output
picResults.Print


'Code option doing sequential searching and defined which case is good.
'Code take information from guest connect information whit information in data and
'search for valid condition.
Found = False
I = 0
Do While ((Not Found) And (I < Ctr1))
I = I + 1
If InputName = PartyName(I) Then Found = True 'sequential searching
Loop
If (Not Found) Then
MsgBox "Wrong name, please try again"
Else
total(I) = Percent(I) * 100
Select Case total(I) 'search for valid condition
Case total(I) = 0 To 3
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is not qualify for Republic of Srpska Assembly, won 0 chairs"
Case total(I) < 3 To 5
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 1 chair"
Case total(I) < 5 To 10
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 2 chairs"
Case total(I) < 10 To 15
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 3 chairs"
Case total(I) < 15 To 20
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 4 chairs"
Case total(I) < 20 To 25
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 5 chairs"
Case total(I) < 25 To 30
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 6 chairs"
Case total(I) < 30 To 35
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 7 chairs"
Case total(I) < 35 To 40
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 8 chairs"
Case total(I) < 40 To 45
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 9 chairs"
Case total(I) < 45 To 50
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 10 chairs"
Case total(I) < 50 To 55
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 11 chairs"
Case total(I) < 55 To 60
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 12 chairs"
Case total(I) < 60 To 65
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 13 chairs"
Case total(I) < 65 To 70
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 14 chairs"
Case total(I) < 70 To 75
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 15 chairs"
Case total(I) < 75 To 80
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 16 chairs"
Case total(I) < 80 To 85
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 17 chairs"
Case total(I) < 85 To 90
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 18 chairs"
Case total(I) < 90 To 95
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 19 chairs"
Case total(I) < 95 To 100
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 20 chairs"
End Select
End If

Close #1 'close input
End Sub

Private Sub cmdMN2_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 2"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County2.txt" For Input As #1

InputName = txtEnterName.Text

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr2 = 0
Sum = 0
Do While Not EOF(1)
Ctr2 = Ctr2 + 1
Input #1, PartyName(Ctr2), Votes(Ctr2)
Sum = Sum + Votes(Ctr2)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr2
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

'Option make empti space in output
picResults.Print


'Code option doing sequential searching and defined which case is good.
'Code take information from guest connect information whit information in data and
'search for valid condition.
Found = False
I = 0
Do While ((Not Found) And (I < Ctr2))
I = I + 1
If InputName = PartyName(I) Then Found = True 'sequential searching
Loop
If (Not Found) Then
MsgBox "Wrong name, please try again"
Else
total(I) = Percent(I) * 100
Select Case total(I) 'search for valid condition
Case total(I) = 0 To 3
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is not qualify for Republic of Srpska Assembly, won 0 chairs"
Case total(I) < 3 To 5
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 1 chair"
Case total(I) < 5 To 10
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 2 chairs"
Case total(I) < 10 To 15
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 3 chairs"
Case total(I) < 15 To 20
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 4 chairs"
Case total(I) < 20 To 25
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 5 chairs"
Case total(I) < 25 To 30
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 6 chairs"
Case total(I) < 30 To 35
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 7 chairs"
Case total(I) < 35 To 40
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 8 chairs"
Case total(I) < 40 To 45
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 9 chairs"
Case total(I) < 45 To 50
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 10 chairs"
Case total(I) < 50 To 55
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 11 chairs"
Case total(I) < 55 To 60
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 12 chairs"
Case total(I) < 60 To 65
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 13 chairs"
Case total(I) < 65 To 70
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 14 chairs"
Case total(I) < 70 To 75
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 15 chairs"
Case total(I) < 75 To 80
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 16 chairs"
Case total(I) < 80 To 85
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 17 chairs"
Case total(I) < 85 To 90
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 18 chairs"
Case total(I) < 90 To 95
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 19 chairs"
Case total(I) < 95 To 100
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 20 chairs"
End Select
End If

Close #1 'close input

End Sub

Private Sub cmdMN3_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 3"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County3.txt" For Input As #1

InputName = txtEnterName.Text

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr3 = 0
Sum = 0
Do While Not EOF(1)
Ctr3 = Ctr3 + 1
Input #1, PartyName(Ctr3), Votes(Ctr3)
Sum = Sum + Votes(Ctr3)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr3
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

'Option make empti space in output
picResults.Print


'Code option doing sequential searching and defined which case is good.
'Code take information from guest connect information whit information in data and
'search for valid condition.
Found = False
I = 0
Do While ((Not Found) And (I < Ctr3))
I = I + 1
If InputName = PartyName(I) Then Found = True 'sequential searching
Loop
If (Not Found) Then
MsgBox "Wrong name, please try again"
Else
total(I) = Percent(I) * 100

Select Case total(I) 'search for valid condition
Case total(I) = 0 To 3
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is not qualify for Republic of Srpska Assembly, won 0 chairs"
Case total(I) < 3 To 5
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 1 chair"
Case total(I) < 5 To 10
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 2 chairs"
Case total(I) < 10 To 15
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 3 chairs"
Case total(I) < 15 To 20
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 4 chairs"
Case total(I) < 20 To 25
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 5 chairs"
Case total(I) < 25 To 30
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 6 chairs"
Case total(I) < 30 To 35
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 7 chairs"
Case total(I) < 35 To 40
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 8 chairs"
Case total(I) < 40 To 45
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 9 chairs"
Case total(I) < 45 To 50
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 10 chairs"
Case total(I) < 50 To 55
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 11 chairs"
Case total(I) < 55 To 60
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 12 chairs"
Case total(I) < 60 To 65
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 13 chairs"
Case total(I) < 65 To 70
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 14 chairs"
Case total(I) < 70 To 75
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 15 chairs"
Case total(I) < 75 To 80
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 16 chairs"
Case total(I) < 80 To 85
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 17 chairs"
Case total(I) < 85 To 90
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 18 chairs"
Case total(I) < 90 To 95
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 19 chairs"
Case total(I) < 95 To 100
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 20 chairs"
End Select
End If

Close #1 'close input

End Sub

Private Sub cmdMN4_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 4"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County4.txt" For Input As #1

InputName = txtEnterName.Text

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr4 = 0
Sum = 0
Do While Not EOF(1)
Ctr4 = Ctr4 + 1
Input #1, PartyName(Ctr4), Votes(Ctr4)
Sum = Sum + Votes(Ctr4)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr4
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

'Option make empti space in output
picResults.Print


'Code option doing sequential searching and defined which case is good.
'Code take information from guest connect information whit information in data and
'search for valid condition.
Found = False
I = 0
Do While ((Not Found) And (I < Ctr4))
I = I + 1
If InputName = PartyName(I) Then Found = True 'sequential searching
Loop
If (Not Found) Then
MsgBox "Wrong name, please try again"
Else
total(I) = Percent(I) * 100

Select Case total(I) 'search for valid condition
Case total(I) = 0 To 3
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is not qualify for Republic of Srpska Assembly, won 0 chairs"
Case total(I) < 3 To 5
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 1 chair"
Case total(I) < 5 To 10
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 2 chairs"
Case total(I) < 10 To 15
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 3 chairs"
Case total(I) < 15 To 20
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 4 chairs"
Case total(I) < 20 To 25
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 5 chairs"
Case total(I) < 25 To 30
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 6 chairs"
Case total(I) < 30 To 35
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 7 chairs"
Case total(I) < 35 To 40
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 8 chairs"
Case total(I) < 40 To 45
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 9 chairs"
Case total(I) < 45 To 50
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 10 chairs"
Case total(I) < 50 To 55
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 11 chairs"
Case total(I) < 55 To 60
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 12 chairs"
Case total(I) < 60 To 65
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 13 chairs"
Case total(I) < 65 To 70
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 14 chairs"
Case total(I) < 70 To 75
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 15 chairs"
Case total(I) < 75 To 80
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 16 chairs"
Case total(I) < 80 To 85
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 17 chairs"
Case total(I) < 85 To 90
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 18 chairs"
Case total(I) < 90 To 95
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 19 chairs"
Case total(I) < 95 To 100
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 20 chairs"
End Select
End If

Close #1 'close input

End Sub

Private Sub cmdMN5_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 5"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County5.txt" For Input As #1

InputName = txtEnterName.Text

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr5 = 0
Sum = 0
Do While Not EOF(1)
Ctr5 = Ctr5 + 1
Input #1, PartyName(Ctr5), Votes(Ctr5)
Sum = Sum + Votes(Ctr5)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr5
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

'Option make empti space in output
picResults.Print


'Code option doing sequential searching and defined which case is good.
'Code take information from guest connect information whit information in data and
'search for valid condition.
Found = False
I = 0
Do While ((Not Found) And (I < Ctr5))
I = I + 1
If InputName = PartyName(I) Then Found = True 'sequential searching
Loop
If (Not Found) Then
MsgBox "Wrong name, please try again"
Else
total(I) = Percent(I) * 100

Select Case total(I) 'search for valid condition
Case total(I) = 0 To 3
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is not qualify for Republic of Srpska Assembly, won 0 chairs"
Case total(I) < 3 To 5
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 1 chair"
Case total(I) < 5 To 10
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 2 chairs"
Case total(I) < 10 To 15
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 3 chairs"
Case total(I) < 15 To 20
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 4 chairs"
Case total(I) < 20 To 25
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 5 chairs"
Case total(I) < 25 To 30
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 6 chairs"
Case total(I) < 30 To 35
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 7 chairs"
Case total(I) < 35 To 40
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 8 chairs"
Case total(I) < 40 To 45
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 9 chairs"
Case total(I) < 45 To 50
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 10 chairs"
Case total(I) < 50 To 55
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 11 chairs"
Case total(I) < 55 To 60
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 12 chairs"
Case total(I) < 60 To 65
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 13 chairs"
Case total(I) < 65 To 70
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 14 chairs"
Case total(I) < 70 To 75
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 15 chairs"
Case total(I) < 75 To 80
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 16 chairs"
Case total(I) < 80 To 85
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 17 chairs"
Case total(I) < 85 To 90
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 18 chairs"
Case total(I) < 90 To 95
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 19 chairs"
Case total(I) < 95 To 100
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 20 chairs"
End Select
End If

Close #1 'close input

End Sub

Private Sub cmdMN6_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 6"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County6.txt" For Input As #1

InputName = txtEnterName.Text

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr6 = 0
Sum = 0
Do While Not EOF(1)
Ctr6 = Ctr6 + 1
Input #1, PartyName(Ctr6), Votes(Ctr6)
Sum = Sum + Votes(Ctr6)
Loop

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr6
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

'Option make empti space in output
picResults.Print


'Code option doing sequential searching and defined which case is good.
'Code take information from guest connect information whit information in data and
'search for valid condition.
Found = False
I = 0
Do While ((Not Found) And (I < Ctr6))
I = I + 1
If InputName = PartyName(I) Then Found = True 'sequential searching
Loop
If (Not Found) Then
MsgBox "Wrong name, please try again"
Else
total(I) = Percent(I) * 100

Select Case total(I) 'search for valid condition
Case total(I) = 0 To 3
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is not qualify for Republic of Srpska Assembly, won 0 chairs"
Case total(I) < 3 To 5
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 1 chair"
Case total(I) < 5 To 10
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 2 chairs"
Case total(I) < 10 To 15
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 3 chairs"
Case total(I) < 15 To 20
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 4 chairs"
Case total(I) < 20 To 25
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 5 chairs"
Case total(I) < 25 To 30
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 6 chairs"
Case total(I) < 30 To 35
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 7 chairs"
Case total(I) < 35 To 40
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 8 chairs"
Case total(I) < 40 To 45
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 9 chairs"
Case total(I) < 45 To 50
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 10 chairs"
Case total(I) < 50 To 55
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 11 chairs"
Case total(I) < 55 To 60
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 12 chairs"
Case total(I) < 60 To 65
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 13 chairs"
Case total(I) < 65 To 70
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 14 chairs"
Case total(I) < 70 To 75
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 15 chairs"
Case total(I) < 75 To 80
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 16 chairs"
Case total(I) < 80 To 85
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 17 chairs"
Case total(I) < 85 To 90
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 18 chairs"
Case total(I) < 90 To 95
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 19 chairs"
Case total(I) < 95 To 100
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes."
picResults.Print Tab(3); "This means party is qualify for Republic of Srpska Assembly, won 20 chairs"
End Select
End If

Close #1 'close input

End Sub

Private Sub cmdPN1_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 1"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County1.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr1 = 0
Sum = 0
Do While Not EOF(1)
Ctr1 = Ctr1 + 1
Input #1, PartyName(Ctr1), Votes(Ctr1)
Sum = Sum + Votes(Ctr1)
Loop

'Code where program sort parties by name
For Pass = 1 To Ctr1 - 1
    For Pos = 1 To Ctr1 - Pass
        If PartyName(Pos) > PartyName(Pos + 1) Then
            Temp = PartyName(Pos)
            PartyName(Pos) = PartyName(Pos + 1)
            PartyName(Pos + 1) = Temp
        
            Temp1 = Votes(Pos)
            Votes(Pos) = Votes(Pos + 1)
            Votes(Pos + 1) = Temp1
            End If
    Next Pos
Next Pass

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr1
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input

End Sub

Private Sub cmdPN2_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 2"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County2.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr2 = 0
Sum = 0
Do While Not EOF(1)
Ctr2 = Ctr2 + 1
Input #1, PartyName(Ctr2), Votes(Ctr2)
Sum = Sum + Votes(Ctr2)
Loop

'Code where program sort parties by name
For Pass = 1 To Ctr2 - 1
    For Pos = 1 To Ctr2 - Pass
        If PartyName(Pos) > PartyName(Pos + 1) Then
            Temp = PartyName(Pos)
            PartyName(Pos) = PartyName(Pos + 1)
            PartyName(Pos + 1) = Temp
        
            Temp1 = Votes(Pos)
            Votes(Pos) = Votes(Pos + 1)
            Votes(Pos + 1) = Temp1
            End If
    Next Pos
Next Pass

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr2
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input

End Sub

Private Sub cmdPN3_Click()

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 3"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County3.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr3 = 0
Sum = 0
Do While Not EOF(1)
Ctr3 = Ctr3 + 1
Input #1, PartyName(Ctr3), Votes(Ctr3)
Sum = Sum + Votes(Ctr3)
Loop

'Code where program sort parties by name
For Pass = 1 To Ctr3 - 1
    For Pos = 1 To Ctr3 - Pass
        If PartyName(Pos) > PartyName(Pos + 1) Then
            Temp = PartyName(Pos)
            PartyName(Pos) = PartyName(Pos + 1)
            PartyName(Pos + 1) = Temp
        
            Temp1 = Votes(Pos)
            Votes(Pos) = Votes(Pos + 1)
            Votes(Pos + 1) = Temp1
            End If
    Next Pos
Next Pass

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr3
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input

End Sub

Private Sub cmdPN4_Click()

picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 4"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County4.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr4 = 0
Sum = 0
Do While Not EOF(1)
Ctr4 = Ctr4 + 1
Input #1, PartyName(Ctr4), Votes(Ctr4)
Sum = Sum + Votes(Ctr4)
Loop

'Code where program sort parties by name
For Pass = 1 To Ctr4 - 1
    For Pos = 1 To Ctr4 - Pass
        If PartyName(Pos) > PartyName(Pos + 1) Then
            Temp = PartyName(Pos)
            PartyName(Pos) = PartyName(Pos + 1)
            PartyName(Pos + 1) = Temp
        
            Temp1 = Votes(Pos)
            Votes(Pos) = Votes(Pos + 1)
            Votes(Pos + 1) = Temp1
            End If
    Next Pos
Next Pass

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr4
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input

End Sub

Private Sub cmdPN5_Click()

picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 5"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County5.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr5 = 0
Sum = 0
Do While Not EOF(1)
Ctr5 = Ctr5 + 1
Input #1, PartyName(Ctr5), Votes(Ctr5)
Sum = Sum + Votes(Ctr5)
Loop

'Code where program sort parties by name
For Pass = 1 To Ctr5 - 1
    For Pos = 1 To Ctr5 - Pass
        If PartyName(Pos) > PartyName(Pos + 1) Then
            Temp = PartyName(Pos)
            PartyName(Pos) = PartyName(Pos + 1)
            PartyName(Pos + 1) = Temp
        
            Temp1 = Votes(Pos)
            Votes(Pos) = Votes(Pos + 1)
            Votes(Pos + 1) = Temp1
            End If
    Next Pos
Next Pass

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr5
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input


End Sub

Private Sub cmdPN6_Click()

picResults.Cls
picResults.Print
picResults.Print Tab(3); "RESULTS FOOR COUNTY 6"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\County6.txt" For Input As #1

'This code option organized array and read Data file
'In same time calculate sum of all votes
Ctr6 = 0
Sum = 0
Do While Not EOF(1)
Ctr6 = Ctr6 + 1
Input #1, PartyName(Ctr6), Votes(Ctr6)
Sum = Sum + Votes(Ctr6)
Loop

'Code where program sort parties by name
For Pass = 1 To Ctr6 - 1
    For Pos = 1 To Ctr6 - Pass
        If PartyName(Pos) > PartyName(Pos + 1) Then
            Temp = PartyName(Pos)
            PartyName(Pos) = PartyName(Pos + 1)
            PartyName(Pos + 1) = Temp
        
            Temp1 = Votes(Pos)
            Votes(Pos) = Votes(Pos + 1)
            Votes(Pos + 1) = Temp1
            End If
    Next Pos
Next Pass

'Option For calculate percent of votes for every party and print results for every party
For I = 1 To Ctr6
Percent(I) = Votes(I) / Sum
picResults.Print Tab(3); PartyName(I); Tab(12); "won"; Votes(I); "votes, or  "; FormatPercent(Percent(I), 2)
Next I

Close #1 'close input


End Sub
