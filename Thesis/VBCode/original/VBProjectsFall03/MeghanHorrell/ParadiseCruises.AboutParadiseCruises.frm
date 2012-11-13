VERSION 5.00
Begin VB.Form frmAboutParadiseCruises 
   BackColor       =   &H00FF0000&
   Caption         =   "About Paradise Cruises"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturnToHome 
      BackColor       =   &H00FF00FF&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   12360
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   12360
      Width           =   2295
   End
   Begin VB.Label lblParadiseCruises 
      BackColor       =   &H00FF0000&
      Caption         =   "Paradise Cruises"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   8175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   $"ParadiseCruises.AboutParadiseCruises.frx":0000
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   2640
      TabIndex        =   0
      Top             =   4200
      Width           =   9735
   End
End
Attribute VB_Name = "frmAboutParadiseCruises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturnToHome_Click()
    frmAboutParadiseCruises.Hide
    frmHome.Show
End Sub
