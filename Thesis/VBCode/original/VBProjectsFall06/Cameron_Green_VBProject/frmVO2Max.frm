VERSION 5.00
Begin VB.Form frmVO2Max 
   BackColor       =   &H00008000&
   Caption         =   "How to Find your VO2 Max"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Homepage"
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdImprovement 
      BackColor       =   &H000000FF&
      Caption         =   "How to Improve your VO2 Max"
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H000000FF&
      Caption         =   "Calculate your VO2 Max"
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdInfoMax 
      BackColor       =   &H000000FF&
      Caption         =   "What is VO2 Max?"
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   3225
      Left            =   6840
      Picture         =   "frmVO2Max.frx":0000
      Top             =   240
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   2280
      Left            =   3240
      Picture         =   "frmVO2Max.frx":3129
      Top             =   360
      Width           =   2700
   End
End
Attribute VB_Name = "frmVO2Max"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to homepage'
Private Sub cmdBack_Click()
    frmVO2Max.Hide
    frmIntroCC.Show
End Sub
'changes from VO2 max page to calculating VO2 max page'
Private Sub cmdCalculate_Click()
    frmVO2Max.Hide
    frmCalculate.Show
End Sub
'changes from VO2 max page to tips to improve VO2 max page'
Private Sub cmdImprovement_Click()
    frmVO2Max.Hide
    frmImprove.Show
End Sub
'changes from VO2 max page to VO2 max information page'
Private Sub cmdInfoMax_Click()
    frmVO2Max.Hide
    frmInfoMax.Show
End Sub
