VERSION 5.00
Begin VB.Form Pepe 
   Caption         =   "P"
   ClientHeight    =   11610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form6"
   ScaleHeight     =   11610
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "back to stats"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   9000
      Width           =   1935
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
      Height          =   1095
      Left            =   8520
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
   Begin VB.PictureBox picResults 
      Height          =   9255
      Left            =   5400
      ScaleHeight     =   9195
      ScaleWidth      =   9075
      TabIndex        =   1
      Top             =   2400
      Width           =   9135
   End
   Begin VB.PictureBox picPicture 
      Height          =   3975
      Left            =   120
      Picture         =   "Pepe.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "Pepe"
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
'Declare variables
Dim CTR As Integer, List(1 To 100) As String
Dim LigaEspanola(1 To 100) As Integer
Dim ChampionsLeague(1 To 100) As Integer
PicResults.Cls
'print the players' stats
PicResults.Print "Name: Pepe"
PicResults.Print "Place of Birth: Maceió, Brazil "
PicResults.Print "Date of Birth:  26/02/1986  "
PicResults.Print "Position: Defender "
PicResults.Print "Weight: 81 kg "
PicResults.Print "Height: 187.1 cm "
PicResults.Print "Nationality: Portuguese/Brazilian"
PicResults.Print "", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"
PicResults.Print "******************************************************************************************************************************"

CTR = CTR + 1
'print the players' stats
PicResults.Print "Games"; Tab(30); "5"; Tab(85); "3"
PicResults.Print ""
PicResults.Print "Home Matches Played"; Tab(30); "3"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Matches Started"; Tab(30); "5"; Tab(85); "3"
PicResults.Print ""
PicResults.Print "Matches Subbed In"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Matches Subbed Out"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Minutes Played"; Tab(30); "437"; Tab(85); "281"
PicResults.Print ""
PicResults.Print "Minutes Played Starter"; Tab(30); "437"; Tab(85); "281"
PicResults.Print ""
PicResults.Print "Goals Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Into Big Box"; Tab(30); "2"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Into Small Box"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Shot on Goal"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Fouls Committed"; Tab(30); "5"; Tab(85); "2"
PicResults.Print ""
PicResults.Print "Fouls Drawn"; Tab(30); "7"; Tab(85); "8"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "1"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Total Passes"; Tab(30); "208"; Tab(85); "151"
PicResults.Print ""
PicResults.Print "Passes incomplete"; Tab(30); "28"; Tab(85); "12"
PicResults.Print ""
PicResults.Print "Passes Completed"; Tab(30); "180"; Tab(85); "139"
End Sub
