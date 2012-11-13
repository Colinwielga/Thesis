VERSION 5.00
Begin VB.Form frmPassingStats 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Passing Stats"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBledsoe 
      Height          =   3255
      Left            =   3240
      Picture         =   "frmPassingStats.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   2040
      Width           =   3855
   End
   Begin VB.PictureBox picPassingStats 
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   9315
      TabIndex        =   0
      Top             =   240
      Width           =   9375
   End
   Begin VB.Label lblBledsoe 
      BackColor       =   &H00C0C0C0&
      Caption         =   "        #11              Drew Bledsoe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   2775
   End
End
Attribute VB_Name = "frmPassingStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub picPassingStats_Paint()
    Dim Player As String
    Dim Attemps As Double, Completions As Double, Yards As Double, TouchDowns As Double, Interceptions As Double, Longest As Double, QBRating As Double
    Open App.Path & "\passingstats.txt" For Input As #1
    Input #1, Player, Attemps, Completions, Yards, TouchDowns, Interceptions, Longest, QBRating
    picPassingStats.Print "Player"; Tab(17); "Attemps"; Tab(28); "Completions"; Tab(43); "Yards"; Tab(52); "Touch Downs"; Tab(70); "Interceptions"; Tab(87); "Longest", "QB Rating"
    picPassingStats.Print "****************************************************************************************************************************************"
    picPassingStats.Print Player; Tab(17); Attemps; Tab(30); Completions; Tab(43); Yards; Tab(56); TouchDowns; Tab(75); Interceptions; Tab(89); Longest, QBRating
    Close #1
End Sub
