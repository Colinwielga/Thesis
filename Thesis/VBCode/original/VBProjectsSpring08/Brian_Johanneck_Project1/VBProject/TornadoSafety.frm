VERSION 5.00
Begin VB.Form TornadoSafety 
   BackColor       =   &H00C0C000&
   Caption         =   "Tornado Safety"
   ClientHeight    =   8280
   ClientLeft      =   2385
   ClientTop       =   1605
   ClientWidth     =   10200
   LinkTopic       =   "Form4"
   Picture         =   "TornadoSafety.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   10200
   Begin VB.CommandButton Command3 
      Caption         =   "Click here to take a look at a picture of a tornado."
      Height          =   1335
      Left            =   8280
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.PictureBox picresults2 
      Height          =   3255
      Left            =   1920
      ScaleHeight     =   3195
      ScaleWidth      =   4755
      TabIndex        =   4
      Top             =   4680
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Learn about some important tornado safety terms."
      Height          =   1815
      Left            =   8400
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.PictureBox picresults 
      Height          =   3855
      Left            =   840
      ScaleHeight     =   3795
      ScaleWidth      =   6795
      TabIndex        =   2
      Top             =   0
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Learn about what to do in case of a tornado."
      Height          =   1575
      Left            =   8400
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Back 
      Caption         =   "Go back to main menu."
      Height          =   1815
      Left            =   8280
      TabIndex        =   0
      Top             =   6120
      Width           =   2175
   End
End
Attribute VB_Name = "TornadoSafety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Back_Click()
TornadoSafety.Hide
Form1.Show
End Sub

Private Sub Command1_Click()
Dim ctr As Integer
Dim J As Integer
Dim Advice1(1 To 20) As String
ctr = 0
picresults.Cls
Open App.Path & "/tornadosafety.txt" For Input As #1
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
Dim Advice1(1 To 20) As String
ctr = 0
picresults.Cls
Open App.Path & "/tornadoterms.txt" For Input As #2
Do While Not EOF(2)
ctr = ctr + 1
Input #2, Advice1(ctr)
Loop
For J = 1 To ctr
picresults.Print Advice1(J)
Next J
Close #2
End Sub

Private Sub Command3_Click()
picresults2.Picture = LoadPicture(App.Path & "/tornado picture.gif")
End Sub
