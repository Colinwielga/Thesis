VERSION 5.00
Begin VB.Form Benzema 
   Caption         =   "KB"
   ClientHeight    =   12570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14625
   LinkTopic       =   "Form4"
   ScaleHeight     =   12570
   ScaleWidth      =   14625
   StartUpPosition =   3  'Windows Default
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
      Height          =   1215
      Left            =   6600
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin VB.PictureBox picPicture 
      Height          =   3735
      Left            =   360
      Picture         =   "Benzema.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   11040
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   10335
      Left            =   4920
      ScaleHeight     =   10275
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   2160
      Width           =   9135
   End
End
Attribute VB_Name = "Benzema"
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




Private Sub cmdInfo_Click()
'Declare Variables
Dim CTR As Integer, List(1 To 100) As String
Dim LigaEspanola(1 To 100) As Integer
Dim ChampionsLeague(1 To 100) As Integer
PicResults.Cls
'print the players' stats
PicResults.Print "Name: Karim Benzema"
PicResults.Print "Place of Birth: Lyon, France "
PicResults.Print "Date of Birth: 19/12/1987 "
PicResults.Print "Position: Forward "
PicResults.Print "Weight:  83.5 kg  "
PicResults.Print "Height: 186.5 cm "
PicResults.Print "Nationality: French"
PicResults.Print "", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"
PicResults.Print "******************************************************************************************************************************"

CTR = CTR + 1
PicResults.Print "Games"; Tab(30); "2"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Away Matches Played"; Tab(30); "4"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Matches Started"; Tab(30); "1"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Matches Subbed In"; Tab(30); "6"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Matches Subbed Out"; Tab(30); "6"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Minutes Played"; Tab(30); "177"; Tab(85); "72"
PicResults.Print ""
PicResults.Print "Minutes Played Starter"; Tab(30); "71"; Tab(85); "72"
PicResults.Print ""
PicResults.Print "Goals Scored"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Headers Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Penalty Kicks Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Free Kicks Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Goals w/ Right Foot"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Goals w/ leftt Foot"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "12"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Fouls Committed"; Tab(30); "10"; Tab(85); "5"
PicResults.Print ""
PicResults.Print "Fouls Drawn"; Tab(30); "7"; Tab(85); "8"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "5"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Total Passes"; Tab(30); "535"; Tab(85); "316"
PicResults.Print ""
PicResults.Print "Passes incomplete"; Tab(30); "92"; Tab(85); "48"
PicResults.Print ""
PicResults.Print "Passes Completed"; Tab(30); "443"; Tab(85); "246"
'picResults.Print List(CTR); Tab(10); LigaEspanola(CTR); Tab(20); ChampionsLeague(CTR)
    
'Loop
End Sub

Private Sub picPicture_Click()
picPicture.Picture = LoadPicture("M:\CS130\Images\BenzemaPic.jpg")
End Sub
