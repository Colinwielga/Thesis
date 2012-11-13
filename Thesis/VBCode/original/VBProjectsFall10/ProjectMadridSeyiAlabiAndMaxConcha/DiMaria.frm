VERSION 5.00
Begin VB.Form DiMaria 
   Caption         =   "ADM"
   ClientHeight    =   12945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   LinkTopic       =   "Form2"
   ScaleHeight     =   12945
   ScaleWidth      =   15360
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
      Height          =   1575
      Left            =   1320
      TabIndex        =   3
      Top             =   10920
      Width           =   2175
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
      Left            =   7080
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.PictureBox picResults 
      Height          =   9615
      Left            =   4920
      ScaleHeight     =   9555
      ScaleWidth      =   8355
      TabIndex        =   1
      Top             =   3360
      Width           =   8415
   End
   Begin VB.PictureBox picPicture 
      Height          =   4215
      Left            =   720
      Picture         =   "DiMaria.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
End
Attribute VB_Name = "DiMaria"
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
'declare Variables
Dim CTR As Integer, List(1 To 100) As String
Dim LigaEspanola(1 To 100) As Integer
Dim ChampionsLeague(1 To 100) As Integer
PicResults.Cls
'print the players' stats
PicResults.Print "Name: Angel Di Maria"
PicResults.Print "Place of Birth: Rosario, Argentina "
PicResults.Print "Date of Birth: 14/02/1988 "
PicResults.Print "Position: Forward "
PicResults.Print "Weight: 75 kg  "
PicResults.Print "Height:  180 cm  "
PicResults.Print "Nationality: Argentine"
PicResults.Print "", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"
PicResults.Print "******************************************************************************************************************************"

CTR = CTR + 1
''print the players' stats
PicResults.Print "Games"; Tab(30); "7"; Tab(85); "3"
PicResults.Print ""
PicResults.Print "Home Matches Played"; Tab(30); "3"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Away Matches Played"; Tab(30); "4"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Matches Started"; Tab(30); "6"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Matches Subbed In"; Tab(30); "1"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Matches Subbed Out"; Tab(30); "6"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Minutes Played"; Tab(30); "419"; Tab(85); "189"
PicResults.Print ""
PicResults.Print "Minutes Played Starter"; Tab(30); "418"; Tab(85); "168"
PicResults.Print ""
PicResults.Print "Goals Scored"; Tab(30); "2"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Headers Scored"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Penalty Kicks Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Free Kicks Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Goals w/ Right Foot"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Goals w/ leftt Foot"; Tab(30); "0"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "13"; Tab(85); "14"
PicResults.Print ""
PicResults.Print "Fouls Committed"; Tab(30); "10"; Tab(85); "5"
PicResults.Print ""
PicResults.Print "Fouls Drawn"; Tab(30); "7"; Tab(85); "8"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "5"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Total Passes"; Tab(30); "224"; Tab(85); "99"
PicResults.Print ""
PicResults.Print "Passes incomplete"; Tab(30); "70"; Tab(85); "23"
PicResults.Print ""
PicResults.Print "Passes Completed"; Tab(30); "156"; Tab(85); "76"
End Sub
