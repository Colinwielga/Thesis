VERSION 5.00
Begin VB.Form Carvalho 
   Caption         =   "RC"
   ClientHeight    =   12645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14655
   LinkTopic       =   "Form7"
   ScaleHeight     =   12645
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   9855
      Left            =   4560
      ScaleHeight     =   9795
      ScaleWidth      =   9435
      TabIndex        =   3
      Top             =   2760
      Width           =   9495
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
      Height          =   1215
      Left            =   6120
      TabIndex        =   2
      Top             =   720
      Width           =   3495
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
      Left            =   1200
      TabIndex        =   1
      Top             =   11160
      Width           =   1815
   End
   Begin VB.PictureBox picPicture 
      Height          =   4215
      Left            =   240
      Picture         =   "Carvalho.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Carvalho"
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
'declare variables
Dim CTR As Integer, List(1 To 100) As String
Dim LigaEspanola(1 To 100) As Integer
Dim ChampionsLeague(1 To 100) As Integer
PicResults.Cls
'print the players' stats
PicResults.Print "Name: Ricardo Carvalho "
PicResults.Print "Place of Birth:  Amarante, Portugal  "
PicResults.Print "Date of Birth:  18/05/1978  "
PicResults.Print "Position: Defender "
PicResults.Print "Weight: 78 kg "
PicResults.Print "Height: 181 cm "
PicResults.Print "Nationality: Portuguese"
PicResults.Print "", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"
PicResults.Print "******************************************************************************************************************************"

CTR = CTR + 1
'print the players' stats
PicResults.Print "Games"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Home Matches Played"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Matches Started"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Matches Subbed In"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Matches Subbed Out"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Minutes Played"; Tab(30); "96"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Minutes Played Starter"; Tab(30); "96"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Goals Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Into Big Box"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Into Small Box"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Shot on Goal"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Fouls Committed"; Tab(30); "10"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Fouls Drawn"; Tab(30); "7"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Total Passes"; Tab(30); "46"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Passes incomplete"; Tab(30); "4"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Passes Completed"; Tab(30); "42"; Tab(85); "0"
End Sub
Private Sub picPicture_Click()
picPicture.Picture = LoadPicture("M:\CS130\Images\RCarvalhoPic.jpg")
End Sub
