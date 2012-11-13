VERSION 5.00
Begin VB.Form frmChinese 
   BackColor       =   &H0080C0FF&
   Caption         =   "Chinese"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmChinese.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmChinese.frx":08CA
   ScaleHeight     =   9150
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCNext 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Click here to see what ingredients you need for Sweet and Sour Pork"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   2895
   End
   Begin VB.PictureBox picCResult 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   3720
      Picture         =   "frmChinese.frx":266D0C
      ScaleHeight     =   3915
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   5040
      Width           =   6255
   End
   Begin VB.CommandButton cmdCReturn 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdCQuit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdCShow 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Show Procedures"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sweet and Sour Pork"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   3
      Top             =   3960
      Width           =   7815
   End
   Begin VB.Image imgChinese 
      Height          =   4320
      Left            =   360
      Picture         =   "frmChinese.frx":2F853E
      Top             =   240
      Width           =   6480
   End
End
Attribute VB_Name = "frmChinese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCNext_Click()
groceryfile = "\Recipes\chineseR.txt "

'Next Step
frmChinese.Hide
frmGroceryStore.Show


End Sub

Private Sub cmdCQuit_Click()
End
End Sub

Private Sub cmdCReturn_Click()

'Return to Homepage
frmCountries.Show
frmChinese.Hide

End Sub

Private Sub cmdCShow_Click()

'Declare Varibles
Dim ChineseCount(1 To 15) As String
Dim CCount As Integer
Dim I As Integer, CTR As Integer

'Open File
Open App.Path & "\Chinese.txt" For Input As #1

CCount = 0

Do Until EOF(1)
    CCount = CCount + 1
    Input #1, ChineseCount(CCount)
Loop

picCResult.Cls
picCResult.Print "Chinese - Sweet and Sour Pork"
picCResult.Print "*****************************************************************"

For I = 1 To CCount
    picCResult.Print ChineseCount(I)
Next I

Close #1

End Sub

