VERSION 5.00
Begin VB.Form Ozil 
   Caption         =   "MO"
   ClientHeight    =   12465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17160
   LinkTopic       =   "Form3"
   ScaleHeight     =   12465
   ScaleWidth      =   17160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
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
      Top             =   10080
      Width           =   2895
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
      Height          =   1455
      Left            =   8040
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
   Begin VB.PictureBox picResults 
      Height          =   9615
      Left            =   6000
      ScaleHeight     =   9555
      ScaleWidth      =   10275
      TabIndex        =   1
      Top             =   2880
      Width           =   10335
   End
   Begin VB.PictureBox picPicture 
      Height          =   4335
      Left            =   600
      Picture         =   "Ozil.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "Ozil"
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
PicResults.Print "Name: Mesut Ozil"
PicResults.Print "Place of Birth: Gelsenkirchen, Germany "
PicResults.Print "Date of Birth: 15/10/1988 "
PicResults.Print "Position:  Midfielder  "
PicResults.Print "Weight: 70 kg. "
PicResults.Print "Height: 181 cm.  "
PicResults.Print "Nationality: German"
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

