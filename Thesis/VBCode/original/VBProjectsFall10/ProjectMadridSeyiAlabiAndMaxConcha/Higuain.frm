VERSION 5.00
Begin VB.Form Higuain 
   Caption         =   "GH"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15675
   LinkTopic       =   "Form10"
   ScaleHeight     =   12210
   ScaleWidth      =   15675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to stats"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1680
      TabIndex        =   3
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Post information on player"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7320
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
   Begin VB.PictureBox picResults 
      Height          =   9015
      Left            =   5520
      ScaleHeight     =   8955
      ScaleWidth      =   9435
      TabIndex        =   1
      Top             =   3240
      Width           =   9495
   End
   Begin VB.PictureBox picPicture 
      Height          =   4215
      Left            =   480
      Picture         =   "Higuain.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "Higuain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
'Hide and/or show desired forms
Statistics.Show
PlayersStat.Hide
OpenPage.Hide
Me.Hide
Information.Hide
Trivia.Hide
End Sub



'Dim Variables
Private Sub cmdInfo_Click()
Dim CTR As Integer, List(1 To 100) As String
Dim LigaEspanola(1 To 100) As Integer
Dim ChampionsLeague(1 To 100) As Integer
PicResults.Cls
'print the players' stats
PicResults.Print "Name: Gonzalo Higuain"
PicResults.Print "Place of Birth: Brest, France  "
PicResults.Print "Date of Birth: 10/12/1987  "
PicResults.Print "Position: Forward "
PicResults.Print "Weight: 81.5 kg  "
PicResults.Print "Height: 184 cm"
PicResults.Print "Nationality: Argentine/French"
PicResults.Print "", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"
PicResults.Print "******************************************************************************************************************************"

CTR = CTR + 1
'print the players' stats
PicResults.Print "Games"; Tab(30); "7"; Tab(85); "3"
PicResults.Print ""
PicResults.Print "Home Matches Played"; Tab(30); "3"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Away Matches Played"; Tab(30); "4"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Matches Started"; Tab(30); "7"; Tab(85); "3"
PicResults.Print ""
PicResults.Print "Matches Subbed In"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Matches Subbed Out"; Tab(30); "4"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Minutes Played"; Tab(30); "623"; Tab(85); "269"
PicResults.Print ""
PicResults.Print "Minutes Played Starter"; Tab(30); "623"; Tab(85); "269"
PicResults.Print ""
PicResults.Print "Goals Scored"; Tab(30); "4"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Headers Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Penalty Kicks Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Free Kicks Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Goals w/ Right Foot"; Tab(30); "3"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Goals w/ leftt Foot"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "22"; Tab(85); "14"
PicResults.Print ""
PicResults.Print "Fouls Committed"; Tab(30); "6"; Tab(85); "5"
PicResults.Print ""
PicResults.Print "Fouls Drawn"; Tab(30); "7"; Tab(85); "8"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "5"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Total Passes"; Tab(30); "190"; Tab(85); "77"
PicResults.Print ""
PicResults.Print "Passes incomplete"; Tab(30); "70"; Tab(85); "23"
PicResults.Print ""
PicResults.Print "Passes Completed"; Tab(30); "156"; Tab(85); "76"
End Sub
