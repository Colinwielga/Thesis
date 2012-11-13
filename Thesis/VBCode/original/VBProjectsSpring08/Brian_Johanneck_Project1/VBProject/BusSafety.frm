VERSION 5.00
Begin VB.Form BusSafety 
   BackColor       =   &H0000FFFF&
   Caption         =   "Bus Safety"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   5295
      Left            =   4440
      Picture         =   "BusSafety.frx":0000
      ScaleHeight     =   5235
      ScaleWidth      =   6315
      TabIndex        =   6
      Top             =   3720
      Width           =   6375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Safety When getting off the bus"
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Safety When Riding the Bus."
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Safety when Getting on the Bus."
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Safety When Walking to the bus."
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox Picresults 
      Height          =   3375
      Left            =   2160
      ScaleHeight     =   3315
      ScaleWidth      =   12075
      TabIndex        =   1
      Top             =   120
      Width           =   12135
   End
   Begin VB.CommandButton Back 
      Caption         =   "Go back to main menu."
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Make sure to stay out odf the danger zone when leaving the bus."
      Height          =   735
      Left            =   2160
      TabIndex        =   7
      Top             =   3840
      Width           =   2175
   End
End
Attribute VB_Name = "BusSafety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Back_Click()
Form1.Show
BusSafety.Hide
End Sub

Private Sub Command1_Click()
Dim ctr As Integer
Dim J As Integer
Dim Advice1(1 To 20) As String
ctr = 0
picresults.Cls
Open App.Path & "/walkingtobusstop.txt" For Input As #1
Do While Not EOF(1)
ctr = ctr + 1
Input #1, Advice1(ctr)
Loop
For J = 1 To ctr
picresults.Print Advice1(J)
Next J
Close #1
End Sub

Private Sub Command2_Click()
Dim ctr As Integer
Dim J As Integer
Dim Advice2(1 To 20) As String
ctr = 0
picresults.Cls
Open App.Path & "/whengettingonthebus.txt" For Input As #2
Do While Not EOF(2)
ctr = ctr + 1
Input #2, Advice2(ctr)
Loop
For J = 1 To ctr
picresults.Print Advice2(J)
Next J
Close #2
End Sub

Private Sub Command3_Click()
Dim ctr As Integer
Dim J As Integer
Dim Advice3(1 To 20) As String
ctr = 0
picresults.Cls
Open App.Path & "/Whenridingthebus.txt" For Input As #3
Do While Not EOF(3)
ctr = ctr + 1
Input #3, Advice3(ctr)
Loop
For J = 1 To ctr
picresults.Print Advice3(J)
Next J
Close #3
End Sub

Private Sub Command4_Click()
Dim ctr As Integer
Dim J As Integer
Dim advice4(1 To 20) As String
ctr = 0
picresults.Cls
Open App.Path & "/whenexitingthebus.txt" For Input As #4
Do While Not EOF(4)
ctr = ctr + 1
Input #4, advice4(ctr)
Loop
For J = 1 To ctr
picresults.Print advice4(J)
Next J
Close #4
End Sub

