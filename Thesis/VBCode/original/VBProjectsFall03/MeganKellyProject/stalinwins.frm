VERSION 5.00
Begin VB.Form stalinwins 
   BackColor       =   &H80000015&
   Caption         =   "Check out my super-cool mustache..."
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form4"
   ScaleHeight     =   6435
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton QuitStalinW 
      BackColor       =   &H80000015&
      Caption         =   "Get me out of this stupid game!!"
      Height          =   975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "Sorry, comrade, but the Man of Steel has bested you.  Don't worry though, he's The People's Dictator.  He'll kill you quickly."
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   3360
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   2370
      Left            =   1800
      Picture         =   "stalinwins.frx":0000
      Top             =   480
      Width           =   2160
   End
End
Attribute VB_Name = "stalinwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub QuitStalinW_Click()
End
End Sub
