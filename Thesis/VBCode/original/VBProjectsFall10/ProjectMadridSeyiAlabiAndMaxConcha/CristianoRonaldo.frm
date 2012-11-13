VERSION 5.00
Begin VB.Form PlayersStat 
   Caption         =   "Player's Stat"
   ClientHeight    =   12495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Post Information on Player"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   3
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Stats"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   840
      TabIndex        =   2
      Top             =   10560
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   9975
      Left            =   4560
      ScaleHeight     =   9915
      ScaleWidth      =   8835
      TabIndex        =   1
      Top             =   2400
      Width           =   8895
   End
   Begin VB.PictureBox picPicture 
      Height          =   3615
      Left            =   360
      Picture         =   "CristianoRonaldo.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "PlayersStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
Statistics.Show
PlayersStat.Hide
OpenPage.Hide
Me.Hide
Information.Hide
Trivia.Hide
End Sub



'dim variables
Private Sub cmdInfo_Click()
Dim CTR As Integer, List(1 To 100) As String
Dim LigaEspanola(1 To 100) As Integer
Dim ChampionsLeague(1 To 100) As Integer
PicResults.Cls
'print the players' stats
PicResults.Print "Name:Cristiano Ronaldo"
PicResults.Print "Place of Birth: Madeira, Portugal "
PicResults.Print "Date of Birth: 05/02/1985 "
PicResults.Print "Position: Forward "
PicResults.Print "Weight: 84.5 kg "
PicResults.Print "Height: 186.5 cm "
PicResults.Print "Nationality: Portuguese"
PicResults.Print "", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"
PicResults.Print "******************************************************************************************************************************"

CTR = CTR + 1
'Show the players stats
PicResults.Print "Games"; Tab(30); "7"; Tab(85); "3"
PicResults.Print ""
PicResults.Print "Home Matches Played"; Tab(30); "4"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Away Matches Played"; Tab(30); "4"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Matches Started"; Tab(30); "7"; Tab(85); "3"
PicResults.Print ""
PicResults.Print "Matches Subbed In"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Matches Subbed Out"; Tab(30); "4"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Minutes Played"; Tab(30); "755"; Tab(85); "281"
PicResults.Print ""
PicResults.Print "Minutes Played Starter"; Tab(30); "755"; Tab(85); "281"
PicResults.Print ""
PicResults.Print "Goals Scored"; Tab(30); "6"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Headers Scored"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Penalty Kicks Scored"; Tab(30); "2"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Free Kicks Scored"; Tab(30); "1"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Goals w/ Right Foot"; Tab(30); "6"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Goals w/ leftt Foot"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "60"; Tab(85); "25"
PicResults.Print ""
PicResults.Print "Fouls Committed"; Tab(30); "6"; Tab(85); "5"
PicResults.Print ""
PicResults.Print "Fouls Drawn"; Tab(30); "7"; Tab(85); "8"
PicResults.Print ""
PicResults.Print "Total Passes"; Tab(30); " 377"; Tab(85); "143"
PicResults.Print ""
PicResults.Print "Passes incomplete"; Tab(30); "66"; Tab(85); "75"
PicResults.Print ""
PicResults.Print "Passes Completed"; Tab(30); "278"; Tab(85); "110"
End Sub

Private Sub Form_Load()
'Dim CTR As Integer, List(1 To 100) As String
'Dim LigaEspanola(1 To 100) As Integer
'Dim ChampionsLeague(1 To 100) As Integer


'Me.Refresh



'If selectedPlayer = 7 Then
   ' lblName.Caption = "Cristiano Ronaldo"
   ' lblNation. = "Portuguese"
    'lblPlace. = "Madeira, Portugal"
    'lblBirth.Caption = "05/02/1985"
    'lblPosition.Caption = "Forward"
   ' lblWeight.Caption = "84.5kg"
    'lblHeight.Caption = "186.5cm"
    
'Open App.Path & "\CristianoRonaldo.txt" For Input As #1
'picResults.Print "List", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"

'Do While EOF(1)
'CTR = CTR + 1
'Input #1, List(CTR), LigaEspanola(CTR), ChampionsLeague(CTR)
'picResults.Print List(CTR); Tab(10); LigaEspanola(CTR); Tab(20); ChampionsLeague(CTR)
'Loop





'picPicture.Image

'ElseIf selectedPlayer = 14 Then
   ' lblName.Caption = "Xabi Alonso"
   ' lblPlace.Caption = "Tolosa , Spain"
   ' lblBirth.Caption = "25/11/1981"
   ' lblPosition.Caption = "Midfielder"
   ' lblWeight.Caption = "79 kg"
   ' lblHeight.Caption = "182.7 cm"
   ' lblNation.Caption = "Spanish"
'picPicture.Image

'ElseIf selectedPlayer = 4 Then
'lblName.Caption = "Sergio Ramos"
'picPicture.Image

'ElseIf selectedPlayer = 2 Then
'lblName.Caption = "Ricardo Carvalho"
'picPicture.Image

'ElseIf selectedPlayer = 3 Then
'lblName.Caption = "Pepe"
'picPicture.Image

'ElseIf selectedPlayer = 9 Then
'lblName.Caption = "Karim Benzema"
'picPicture.Image

'ElseIf selectedPlayer = 12 Then
'lblName.Caption = "Marcelo"
'picPicture.Image

'ElseIf selectedPlayer = 20 Then
'lblName.Caption = "Gonzalo Higuain"
'picPicture.Image

'ElseIf selectedPlayer = 22 Then
'lblName.Caption = "Angel Di Maria"
'picPicture.Image

'ElseIf selectedPlayer = 23 Then
'lblName.Caption = "Mesut Ozil"
'picPicture.Image

'ElseIf selectedPlayer = 1 Then
'lblName.Caption = "Iker Casillas"
'picPicture.Image

'End If

End Sub




Private Sub picPicture_Click()
picPicture.Picture = LoadPicture("M:\CS130\Images\gallery1.jpg")
End Sub

