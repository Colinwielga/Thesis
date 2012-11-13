VERSION 5.00
Begin VB.Form frmChineseVeggie 
   BackColor       =   &H00404040&
   Caption         =   "Chinese Vegetarian"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmChineseVeggie.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmChineseVeggie.frx":08CA
   ScaleHeight     =   9060
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCVNext 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Click here to see what you need for Vegetables fried rice "
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   3255
   End
   Begin VB.CommandButton cmdCVReturn 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton cmdCVQuit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton cmdCVShow 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Show Procedures"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   3255
   End
   Begin VB.PictureBox picCVResult 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5895
      Left            =   3960
      Picture         =   "frmChineseVeggie.frx":24090C
      ScaleHeight     =   5835
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Vegetables fried rice"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   1320
      TabIndex        =   5
      Top             =   360
      Width           =   8535
   End
   Begin VB.Image imgChineseV 
      Height          =   4740
      Left            =   240
      Picture         =   "frmChineseVeggie.frx":30F68E
      Top             =   1560
      Width           =   3630
   End
End
Attribute VB_Name = "frmChineseVeggie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCVNext_Click()

groceryfile = "\Recipes\chineseVR.txt "

'Next Step
frmChineseVeggie.Hide
frmGroceryStore.Show

End Sub

Private Sub cmdCVReturn_Click()

'Return to Homepage
frmCountries.Show
frmChineseVeggie.Hide

End Sub

Private Sub cmdCVShow_Click()

'Declare Varibles
Dim ChineseVCount(1 To 12) As String
Dim CVCount As Integer
Dim I As Integer, CTR As Integer

'Open File
Open App.Path & "\ChineseV.txt" For Input As #1

CVCount = 0

Do Until EOF(1)
    CVCount = CVCount + 1
    Input #1, ChineseVCount(CVCount)
Loop

picCVResult.Cls
picCVResult.Print "Chinese - Vegetables fried rice"
picCVResult.Print "*****************************************************************"

For I = 1 To CVCount
    picCVResult.Print ChineseVCount(I)
Next I

Close #1

End Sub

Private Sub cmdCVQuit_Click()
End
End Sub




