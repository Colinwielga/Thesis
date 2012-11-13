VERSION 5.00
Begin VB.Form Marcelo 
   Caption         =   "M"
   ClientHeight    =   12270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15825
   LinkTopic       =   "Form8"
   ScaleHeight     =   12270
   ScaleWidth      =   15825
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
      Height          =   1335
      Left            =   6960
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
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
      Height          =   1455
      Left            =   1440
      TabIndex        =   2
      Top             =   9360
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   9375
      Left            =   5280
      ScaleHeight     =   9315
      ScaleWidth      =   8595
      TabIndex        =   1
      Top             =   2880
      Width           =   8655
   End
   Begin VB.PictureBox picPicture 
      Height          =   4095
      Left            =   480
      Picture         =   "Marcelo.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
   End
End
Attribute VB_Name = "Marcelo"
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
PicResults.Print "Name: Marcelo"
PicResults.Print "Place of Birth:  Rio de Janeiro, Brazil "
PicResults.Print "Date of Birth:  12/05/1988  "
PicResults.Print "Position: Defender "
PicResults.Print "Weight: 73 kg  "
PicResults.Print "Height:  171.9 cm  "
PicResults.Print "Nationality: Brazilian"
PicResults.Print "", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"
PicResults.Print "******************************************************************************************************************************"

CTR = CTR + 1
'print the players' stats
PicResults.Print "Games"; Tab(30); "7"; Tab(85); "3"
PicResults.Print ""
PicResults.Print "Home Matches Played"; Tab(30); "3"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Matches Started"; Tab(30); "7"; Tab(85); "3"
PicResults.Print ""
PicResults.Print "Matches Subbed In"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Matches Subbed Out"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Minutes Played"; Tab(30); "663"; Tab(85); "281"
PicResults.Print ""
PicResults.Print "Minutes Played Starter"; Tab(30); "663"; Tab(85); "281"
PicResults.Print ""
PicResults.Print "Goals Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Into Big Box"; Tab(30); "14"; Tab(85); "6"
PicResults.Print ""
PicResults.Print "Crosses Into Small Box"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Shot on Goal"; Tab(30); "3"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Fouls Committed"; Tab(30); "7"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Fouls Drawn"; Tab(30); "7"; Tab(85); "8"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "4"; Tab(85); "6"
PicResults.Print ""
PicResults.Print "Total Passes"; Tab(30); "460"; Tab(85); "226"
PicResults.Print ""
PicResults.Print "Passes incomplete"; Tab(30); "116"; Tab(85); "37"
PicResults.Print ""
PicResults.Print "Passes Completed"; Tab(30); "344"; Tab(85); "189"
End Sub
