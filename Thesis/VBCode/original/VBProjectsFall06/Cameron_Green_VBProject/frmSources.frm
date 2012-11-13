VERSION 5.00
Begin VB.Form frmSources 
   BackColor       =   &H00008000&
   Caption         =   "Sources Used"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Homepage"
      Height          =   855
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblResultsPage 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "http://sporting-heroes.net/athletics-heroes/stats_athletics/olympics/olympics.asp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   10095
   End
   Begin VB.Label lblGoogleImages 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "http://images.google.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   10095
   End
   Begin VB.Label lblCalculator 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "http://www.coolrunning.com/engine/4/4_1/94.shtml"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   10215
   End
   Begin VB.Label lblBMIPage 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "http://www.nhlbisupport.com/bmi/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   10095
   End
   Begin VB.Label lblVO2Max 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "http://www.brianmac.demon.co.uk/vo2max.htm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   10095
   End
   Begin VB.Label lblWikipedia 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "www.wikipedia.org"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   9855
   End
End
Attribute VB_Name = "frmSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to homepage from sources used page'
Private Sub cmdBack_Click()
    frmIntroCC.Show
    frmSources.Hide
End Sub
