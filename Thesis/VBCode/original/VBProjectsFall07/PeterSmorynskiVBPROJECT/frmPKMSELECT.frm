VERSION 5.00
Begin VB.Form frmPKMSELECT 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue Game"
      Height          =   735
      Left            =   5160
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Previous Screen"
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Game"
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdBulb 
      Caption         =   "BULBASAUR"
      Height          =   1455
      Left            =   9360
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdSquirt 
      Caption         =   "SQUIRTLE"
      Height          =   1455
      Left            =   6120
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdChar 
      Caption         =   "CHARMANDER"
      Height          =   1455
      Left            =   3240
      TabIndex        =   0
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   2895
      Left            =   9120
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   2895
      Left            =   6000
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   3120
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmPKMSELECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
