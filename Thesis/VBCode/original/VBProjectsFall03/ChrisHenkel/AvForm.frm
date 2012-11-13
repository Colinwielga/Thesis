VERSION 5.00
Begin VB.Form AvForm 
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   600
      TabIndex        =   3
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CommandButton AveTime 
      Caption         =   "Find the Average Time"
      Height          =   1095
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton AvePos 
      Caption         =   "Find the Average Position"
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   5400
      ScaleHeight     =   5715
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "AvForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Skier(1 To 72) As String, Country(1 To 72) As String, Time(1 To 72) As String
Dim Place(1 To 72) As Integer, J As Integer
Private Sub Form_Load()
Open "M:\ChrisHenkel\RaceResults.txt" For Input As #1
For J = 1 To 72
    Input #1, Place(J), Skier(J), Country(J), Time(J)
Next J
Close
End Sub
