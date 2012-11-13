VERSION 5.00
Begin VB.Form Ramos 
   Caption         =   "SR"
   ClientHeight    =   12750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form5"
   ScaleHeight     =   12750
   ScaleWidth      =   14640
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
      Height          =   1695
      Left            =   1080
      TabIndex        =   3
      Top             =   8760
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   10095
      Left            =   4800
      ScaleHeight     =   10035
      ScaleWidth      =   9195
      TabIndex        =   2
      Top             =   2520
      Width           =   9255
   End
   Begin VB.PictureBox picPicture 
      Height          =   4215
      Left            =   240
      Picture         =   "Ramos.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   720
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
      Height          =   1335
      Left            =   6960
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "Ramos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
'declare variables
Statistics.Show
PlayersStat.Hide
OpenPage.Hide
Me.Hide
Information.Hide
Trivia.Hide
End Sub




Private Sub cmdInfo_Click()
Dim CTR As Integer, List(1 To 100) As String
Dim LigaEspanola(1 To 100) As Integer
Dim ChampionsLeague(1 To 100) As Integer
PicResults.Cls
'print the players' stats
PicResults.Print "Name: Sergio Ramos"
PicResults.Print ""
PicResults.Print "Place of Birth: Seville, Spain "
PicResults.Print ""
PicResults.Print "Date of Birth: 30/03/1986 "
PicResults.Print ""
PicResults.Print "Position: Defender "
PicResults.Print ""
PicResults.Print "Weight:  81 kg "
PicResults.Print ""
PicResults.Print "Height: 183.2 cm "
PicResults.Print ""
PicResults.Print "Nationality: Spanish"
PicResults.Print ""
PicResults.Print "", , "Liga Española 1ª División 2010-11", , "Champions League 2010-11"
PicResults.Print "******************************************************************************************************************************"

CTR = CTR + 1
'print the players' stats
PicResults.Print "Games"; Tab(30); "6"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Home Matches Played"; Tab(30); "3"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Matches Started"; Tab(30); "6"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Matches Subbed In"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Matches Subbed Out"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Minutes Played"; Tab(30); "569"; Tab(85); "94"
PicResults.Print ""
PicResults.Print "Minutes Played Starter"; Tab(30); "569"; Tab(85); "94"
PicResults.Print ""
PicResults.Print "Goals Scored"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Into Big Box"; Tab(30); "13"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Into Small Box"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Crosses Shot on Goal"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Fouls Committed"; Tab(30); "11"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Fouls Drawn"; Tab(30); "10"; Tab(85); "8"
PicResults.Print ""
PicResults.Print "Shots"; Tab(30); "6"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Total Passes"; Tab(30); "434"; Tab(85); "81"
PicResults.Print ""
PicResults.Print "Passes incomplete"; Tab(30); "369"; Tab(85); "72"
PicResults.Print ""
PicResults.Print "Passes Completed"; Tab(30); "65"; Tab(85); "9"
End Sub
Private Sub picPicture_Click()
picPicture.Picture = LoadPicture("M:\CS130\Images\RamosPic.jpg")
End Sub
