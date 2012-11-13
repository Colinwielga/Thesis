VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Left            =   3840
      ScaleHeight     =   4995
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
Dim SwimmerName As String
Dim SwimmerTime As Double
Dim SwimmerEvent As String
Dim CTR As Integer

Open App.Path & "\faketimes.txt" For Output As #1

SwimmerName = InputBox("Please Enter Swimmer Name.", "Swimmer.")
SwimmerEvent = InputBox("Please Enter The Event.", "Event.")
SwimmerTime = InputBox("Please Enter Time in Seconds.", "Time")

Write #1, SwimmerName, SwimmerEvent, SwimmerTime

 
Close

End Sub

Private Sub Command2_Click()
Form1.Hide
Form2.Show

End Sub
