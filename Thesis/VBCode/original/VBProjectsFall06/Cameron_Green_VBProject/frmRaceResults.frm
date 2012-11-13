VERSION 5.00
Begin VB.Form frmRaceResults 
   BackColor       =   &H00008000&
   Caption         =   "Past Race Results"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Homepage"
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdOlympics 
      BackColor       =   &H000000FF&
      Caption         =   "Olympic Race Results"
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   3720
      Picture         =   "frmRaceResults.frx":0000
      Top             =   1560
      Width           =   4500
   End
End
Attribute VB_Name = "frmRaceResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this button goes back to homepage from race results page'
Private Sub cmdBack_Click()
    frmRaceResults.Hide
    frmIntroCC.Show
End Sub

'this button goes from the race results page to the olympic results page'
Private Sub cmdOlympics_Click()
    frmRaceResults.Hide
    frmOlympics.Show
End Sub
