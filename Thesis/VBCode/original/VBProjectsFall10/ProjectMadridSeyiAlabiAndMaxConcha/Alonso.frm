VERSION 5.00
Begin VB.Form Alonso 
   Caption         =   "XA"
   ClientHeight    =   12615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14535
   LinkTopic       =   "Form11"
   ScaleHeight     =   12615
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   3735
      Left            =   360
      Picture         =   "Alonso.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   600
      Width           =   4095
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
      Left            =   6360
      TabIndex        =   2
      Top             =   840
      Width           =   4215
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
   Begin VB.PictureBox picResults 
      Height          =   10335
      Left            =   5280
      ScaleHeight     =   10275
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   2280
      Width           =   9015
   End
End
Attribute VB_Name = "Alonso"
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
picResults.Cls
'print the players' stats
picResults.Print "Name: Xabi Alonso"
picResults.Print "Place of Birth: Tolosa, Spain "
picResults.Print "Date of Birth: 25/11/1981 "
picResults.Print "Position: Midfielder "
picResults.Print "Weight: 79 kg "
picResults.Print "Height: 182.7 cm  "
picResults.Print "Nationality: Spanish"
picResults.Print "", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"
picResults.Print "*****************************************************************************************************************************************************"

CTR = CTR + 1

picResults.Print "Games"; Tab(30); "7"; Tab(85); "3"
picResults.Print ""
picResults.Print "Home Matches Played"; Tab(30); "3"; Tab(85); "2"
picResults.Print ""
picResults.Print "Matches Started"; Tab(30); "7"; Tab(85); "3"
picResults.Print ""
picResults.Print "Matches Subbed In"; Tab(30); "0"; Tab(85); "0"
picResults.Print ""
picResults.Print "Matches Subbed Out"; Tab(30); "0"; Tab(85); "1"
picResults.Print ""
picResults.Print "Minutes Played"; Tab(30); "663"; Tab(85); "272"
picResults.Print ""
picResults.Print "Minutes Played Starter"; Tab(30); "663"; Tab(85); "272"
picResults.Print ""
picResults.Print "Goals Scored"; Tab(30); "0"; Tab(85); "0"
picResults.Print ""
picResults.Print "Crosses Into Big Box"; Tab(30); "25"; Tab(85); "15"
picResults.Print ""
picResults.Print "Crosses Into Small Box"; Tab(30); "0"; Tab(85); "0"
picResults.Print ""
picResults.Print "Crosses Shot on Goal"; Tab(30); "4"; Tab(85); "2"
picResults.Print ""
picResults.Print "Fouls Committed"; Tab(30); "10"; Tab(85); "5"
picResults.Print ""
picResults.Print "Fouls Drawn"; Tab(30); "7"; Tab(85); "8"
picResults.Print ""
picResults.Print "Shots"; Tab(30); "5"; Tab(85); "1"
picResults.Print ""
picResults.Print "Total Passes"; Tab(30); "535"; Tab(85); "316"
picResults.Print ""
picResults.Print "Passes incomplete"; Tab(30); "92"; Tab(85); "48"
picResults.Print ""
picResults.Print "Passes Completed"; Tab(30); "443"; Tab(85); "246"
End Sub
Private Sub picPicture_Click()
'load the file containig the picture and show in the picturebox
picPicture.Picture = LoadPicture("M:\CS130\Images\AlonsoPic.jpg")
End Sub
