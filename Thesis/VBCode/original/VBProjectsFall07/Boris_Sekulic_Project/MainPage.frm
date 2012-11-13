VERSION 5.00
Begin VB.Form frmMainPage 
   BackColor       =   &H8000000E&
   Caption         =   "Republic of Srpska, election 2006"
   ClientHeight    =   9930
   ClientLeft      =   5040
   ClientTop       =   855
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   Picture         =   "MainPage.frx":0000
   ScaleHeight     =   9930
   ScaleWidth      =   10440
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   5955
      TabIndex        =   5
      Top             =   2760
      Width           =   6015
   End
   Begin VB.CommandButton cmdMaps 
      BackColor       =   &H00C0E0FF&
      Caption         =   "View map of counties in Republic of Srpska"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   2415
   End
   Begin VB.CommandButton cmdRSInformation 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Other informations about Republic of Srpska"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8880
      Width           =   2415
   End
   Begin VB.CommandButton cmdCountyu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "View results by Counties"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton smdReadData 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Basic information about elections in Republic of Srpska"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2415
   End
End
Attribute VB_Name = "frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCountyu_Click()

'Cod Option that connect two forms
'In this case Main Page and Counties
frmCounties.Show
frmMainPage.Hide

End Sub

Private Sub cmdMaps_Click()

'Cod Option that connect two forms
'In this case Main Page and Maps
frmMaps.Show
frmMainPage.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRSInformation_Click()

'Cod Option that connect two forms
'In this case Main Page and RSInformation
frmRSInformation.Show
frmMainPage.Hide
End Sub

Private Sub smdReadData_Click()

'This code defineds Private Values
Dim Sum1 As Double, Sum2 As Double, Sum3 As Double, Sum4 As Double
Dim Sum5 As Double, Sum6 As Double, Sum7 As Double
Dim PercentVotes As Single, PercentCorectVotes As Double
Dim PercentIncorectVotes As Long, Ctr8 As Integer, TotalVotes As Double

'This code option Open Data files as input
Open App.Path & "\County1.txt" For Input As #1
Open App.Path & "\County2.txt" For Input As #2
Open App.Path & "\County3.txt" For Input As #3
Open App.Path & "\County4.txt" For Input As #4
Open App.Path & "\County5.txt" For Input As #5
Open App.Path & "\County6.txt" For Input As #6
Open App.Path & "\Incorrect votes.txt" For Input As #7
Open App.Path & "\Parties_List.txt" For Input As #8


'This code option organized 6 arraes bellow, read Data file for every county
'as input and in same time calculate sum of all votes for taht county
Ctr1 = 0
Sum1 = 0
Do While Not EOF(1)
Ctr1 = Ctr1 + 1
Input #1, PartyName(Ctr1), Votes(Ctr1)
Sum1 = Sum1 + Votes(Ctr1)
Loop

Ctr2 = 0
Sum2 = 0
Do While Not EOF(2)
Ctr2 = Ctr2 + 1
Input #2, PartyName(Ctr2), Votes(Ctr2)
Sum2 = Sum2 + Votes(Ctr2)
Loop

Ctr3 = 0
Sum3 = 0
Do While Not EOF(3)
Ctr3 = Ctr3 + 1
Input #3, PartyName(Ctr3), Votes(Ctr3)
Sum3 = Sum3 + Votes(Ctr3)
Loop

Ctr4 = 0
Sum4 = 0
Do While Not EOF(4)
Ctr4 = Ctr4 + 1
Input #4, PartyName(Ctr4), Votes(Ctr4)
Sum4 = Sum4 + Votes(Ctr4)
Loop

Ctr5 = 0
Sum5 = 0
Do While Not EOF(5)
Ctr5 = Ctr5 + 1
Input #5, PartyName(Ctr5), Votes(Ctr5)
Sum5 = Sum5 + Votes(Ctr5)
Loop

Ctr6 = 0
Sum6 = 0
Do While Not EOF(6)
Ctr6 = Ctr6 + 1
Input #6, PartyName(Ctr6), Votes(Ctr6)
Sum6 = Sum5 + Votes(Ctr6)
Loop


'This code option organized array read Data file for incorrect votes and
'In same time calculate sum that votes
Ctr7 = 0
Sum7 = 0
Do While Not EOF(7)
Ctr7 = Ctr7 + 1
Input #7, CountyName(Ctr7), Votes(Ctr7)
Sum7 = Sum7 + Votes(Ctr7)
Loop

'Calculate votes for elections and percent of that votes from all registered
TotalVotes = Sum1 + Sum2 + Sum3 + Sum4 + Sum5 + Sum6 + Sum7
PercentVotes = TotalVotes / 1051068

'Calculate sum of all correct votes
TotalCorectVotes = Sum1 + Sum2 + Sum3 + Sum4 + Sum5 + Sum6

PercentCorectVotes = TotalCorectVotes / TotalVotes 'Percent of all correct votes
PercentIncorectVotes = Sum7 / TotalVotes            'Percent of all incorrect votes


'Print information in output
picResults.Print
picResults.Print "Republic of Srpska (RS) is Serbian part of Bosnia and Herzegovina."
picResults.Print "*******************************************************"
picResults.Print "People of RS elect members of House of Representatives "
picResults.Print
picResults.Print "Last election for House of Representatives and for President of RS"
picResults.Print "where implemented on September 26, 2006."
picResults.Print
picResults.Print "Number of registered voters for election 2006 in RS is    1.051.068 "
picResults.Print "Number of voters that voted on last election is"; Tab(48); TotalVotes; Tab(58); FormatPercent(PercentVotes, 2)
picResults.Print "Nuber of correct votes is "; Tab(48); TotalCorectVotes; Tab(58); FormatPercent(PercentCorectVotes, 2)
picResults.Print "Number of incorrect votes is "; Tab(48); Sum7; Tab(58); FormatPercent(PercentIncorectVotes, 2)
picResults.Print
picResults.Print "Parties registered for election 2006 are:"
picResults.Print


'Read data for party list and print all information in output
Ctr8 = 0
Do While Not EOF(8)
Ctr8 = Ctr8 + 1
Input #8, PartyName(Ctr8)
picResults.Print PartyName(Ctr8)
Loop

'Close inputs
Close #1
Close #2
Close #3
Close #4
Close #5
Close #6
Close #7
Close #8

End Sub
