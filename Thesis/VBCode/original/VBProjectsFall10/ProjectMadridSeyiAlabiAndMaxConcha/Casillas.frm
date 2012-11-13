VERSION 5.00
Begin VB.Form Casillas 
   Caption         =   "IC"
   ClientHeight    =   13350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15300
   LinkTopic       =   "Form9"
   ScaleHeight     =   13350
   ScaleWidth      =   15300
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
      Height          =   1455
      Left            =   1560
      TabIndex        =   3
      Top             =   10680
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
      Height          =   1575
      Left            =   7320
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.PictureBox picResults 
      Height          =   10095
      Left            =   5040
      ScaleHeight     =   10035
      ScaleWidth      =   9435
      TabIndex        =   1
      Top             =   3000
      Width           =   9495
   End
   Begin VB.PictureBox picPicture 
      Height          =   3975
      Left            =   600
      Picture         =   "Casillas.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "Casillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
PicResults.Print "Name: Iker Casillas"
PicResults.Print "Place of Birth: Mostoles, Spain "
PicResults.Print "Date of Birth: 20/05/1981 "
PicResults.Print "Position: Keeper "
PicResults.Print "Weight:  85.5 kg "
PicResults.Print "Height:  182.2 cm "
PicResults.Print "Nationality: Spanish"
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
PicResults.Print "Matches Subbed Out"; Tab(30); "0"; Tab(85); "1"
PicResults.Print ""
PicResults.Print "Minutes Played"; Tab(30); "663"; Tab(85); "272"
PicResults.Print ""
PicResults.Print "Minutes Played Starter"; Tab(30); "663"; Tab(85); "272"
PicResults.Print ""
PicResults.Print "Goals Conceded"; Tab(30); "3"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Home Goals Conceded"; Tab(30); "1"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Away Goals Conceded"; Tab(30); "2"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Steals"; Tab(30); "34"; Tab(85); "15"
PicResults.Print ""
PicResults.Print "Fouls Committed"; Tab(30); "0"; Tab(85); "0"
PicResults.Print ""
PicResults.Print "Fouls Drawn"; Tab(30); "7"; Tab(85); "8"
PicResults.Print ""
PicResults.Print "Turnovers"; Tab(30); "35"; Tab(85); "10"
PicResults.Print ""
PicResults.Print "Total Passes"; Tab(30); "127"; Tab(85); "54"
PicResults.Print ""
PicResults.Print "Passes incomplete"; Tab(30); "52"; Tab(85); "12"
PicResults.Print ""
PicResults.Print "Passes Completed"; Tab(30); "75"; Tab(85); "42"

End Sub


