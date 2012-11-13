VERSION 5.00
Begin VB.Form frmCongrats 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Congratulations"
   ClientHeight    =   6870
   ClientLeft      =   8505
   ClientTop       =   1125
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   9585
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Window"
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lblClickToPlay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Double Click to play!"
      Height          =   495
      Left            =   8160
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sound Clip:    Queen - We  Are the Champions"
      Height          =   855
      Left            =   8160
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.OLE OLE1 
      Class           =   "SoundRec"
      Height          =   615
      Left            =   7320
      OleObjectBlob   =   "frmCongrats.frx":0000
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Image imgMedalSilver 
      Height          =   2250
      Left            =   7080
      Picture         =   "frmCongrats.frx":147218
      Top             =   1320
      Width           =   2250
   End
   Begin VB.Image imgMedalGold 
      Height          =   2400
      Left            =   120
      Picture         =   "frmCongrats.frx":14CAE2
      Top             =   1080
      Width           =   2250
   End
   Begin VB.Label lblBobAndGrant 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bob and Grant for a 1st and 2nd place finish!!!"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   600
      TabIndex        =   0
      Top             =   4680
      Width           =   8385
   End
   Begin VB.Image imgCongrats 
      Height          =   4335
      Left            =   2280
      Picture         =   "frmCongrats.frx":14D752
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4740
   End
End
Attribute VB_Name = "frmCongrats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project: SJU_Ski_Team
'Form name: frmCongrats
'Author: Kevin Neal
'Written: March 23, 2009
'Object: 1)Congratulate Bob and Grant for outstanding performances
        '2)Add a sound clip to my project
        '3)More pictures

Private Sub cmdClose_Click()
    'Closes form
    frmCongrats.Hide
End Sub


