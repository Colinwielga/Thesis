VERSION 5.00
Begin VB.Form frmTransportation 
   Caption         =   "Using Public Transportation"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   6240
      TabIndex        =   7
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdGoToHome 
      Caption         =   "Return to Home Page"
      Height          =   975
      Left            =   4200
      TabIndex        =   6
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox txtDestination 
      Height          =   855
      Left            =   6960
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtLocation 
      Height          =   855
      Left            =   2760
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblDestinations 
      Caption         =   "Popular Destinations"
      Height          =   4335
      Left            =   4680
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblDestination 
      Caption         =   "Enter the number corresponding to the popular destination you would like to visit."
      Height          =   855
      Left            =   4560
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label lblLocation 
      Caption         =   "Enter the number corresponding to the station you are nearest to."
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label lblStations 
      Caption         =   "Stations"
      Height          =   4335
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmTransportation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

End Sub
