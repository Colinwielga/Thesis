VERSION 5.00
Begin VB.Form frmFindGames 
   Caption         =   "Look Up Games"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "OCR A Extended"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   3000
      ScaleHeight     =   4275
      ScaleWidth      =   8595
      TabIndex        =   8
      Top             =   360
      Width           =   8655
   End
   Begin VB.CommandButton cmdSortNew 
      Caption         =   "Release"
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortSystem 
      Caption         =   "System"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortTitle 
      Caption         =   "Title"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturnToWant 
      Caption         =   "Return"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdPS3 
      Caption         =   "PS3"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdWii 
      Caption         =   "Wii"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdXbox360 
      Caption         =   "Xbox 360"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdLoadGames 
      Caption         =   "Load games"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblSort 
      Caption         =   "Sort Games By:"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmLookupGame.frx":0000
      Top             =   -360
      Width           =   12000
   End
End
Attribute VB_Name = "frmFindGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdLoadGames_Click()
Dim ctr As Integer
    Open App.Path & "\GamesList.txt" For Input As #1
        ctr = 0
        picResults.Print "Title", , , "Platform", , , "Release"
        picResults.Print "*****************************************************************************************************************************"
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, Title(ctr), Platform(ctr), Release(ctr)
        picResults.Print Title(ctr); Platform(ctr); Release(ctr)
    Loop
    picResults.Print "*****************************************"
    Close #1
End Sub
Private Sub cmdReturnToWant_Click()
    frmFindGames.Hide
    frmSelectWant.Show
End Sub

