VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   5955
      TabIndex        =   9
      Top             =   2160
      Width           =   6015
   End
   Begin VB.PictureBox Picresults 
      BackColor       =   &H008080FF&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   7155
      TabIndex        =   8
      Top             =   7680
      Width           =   7215
   End
   Begin VB.CommandButton order 
      Caption         =   "Click here to get the list of resources in alphebetical order accoding to what safety they are."
      Height          =   1095
      Left            =   9120
      TabIndex        =   7
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton resources 
      Caption         =   "Click here to see the resources used and what  part of safety they corrisponded to."
      Height          =   1095
      Left            =   7320
      TabIndex        =   6
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton END 
      Caption         =   "End program"
      Height          =   1575
      Left            =   7920
      TabIndex        =   4
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton gotoSafetyQuiz 
      Caption         =   "See how much you have learned with a Safety Quz."
      Height          =   1455
      Left            =   7920
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton gotoTornadosafety 
      Caption         =   "Go to take a look at Tornado Safety."
      Height          =   1455
      Left            =   5400
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton gotoFiresafety 
      Caption         =   "Go to take a look at Fire Safety."
      Height          =   1455
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton gotoBussafety 
      BackColor       =   &H80000005&
      Caption         =   "Go to take a look at Bus safety."
      Height          =   1455
      Left            =   600
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Click on any button for information on that topic!!"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim website(1 To 20) As String
Dim safety(1 To 20) As String
Dim ctr As Integer
Dim J As Integer
Private Sub END_Click()
End
End Sub
Private Sub gotoBussafety_Click()
Form1.Hide
BusSafety.Show
End Sub

Private Sub gotoFiresafety_Click()
Form1.Hide
FireSafety.Show
End Sub

Private Sub gotoSafetyQuiz_Click()
Form1.Hide
SafetyQuiz.Show
End Sub

Private Sub gotoTornadosafety_Click()
Form1.Hide
TornadoSafety.Show
End Sub

Private Sub order_Click()
Dim tempplace As String
Dim tempplace2 As String
Dim pos As Integer
Dim pass As Integer
ctr = 0
Open App.Path & "/resources.txt" For Input As #1
Open App.Path & "/websites.txt" For Input As #2
Do While Not EOF(1)
ctr = ctr + 1
Input #1, website(ctr)
Loop
ctr = 0
Do While Not EOF(2)
ctr = ctr + 1
Input #2, safety(ctr)
Loop
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
    If website(pos) > website(pos + 1) Then
    tempplace = safety(pos)
    safety(pos) = safety(pos + 1)
    safety(pos + 1) = tempplace
    tempplace2 = website(pos)
    website(pos) = website(pos + 1)
    website(pos + 1) = tempplace2
    End If
    Next pos
Next pass
Picresults.Cls
For J = 1 To ctr
Picresults.Print safety(J); " is for "; website(J)
Next J
Close #1
Close #2
End Sub

Private Sub resources_Click()
ctr = 0
Open App.Path & "/resources.txt" For Input As #1
Open App.Path & "/websites.txt" For Input As #2
Do While Not EOF(1)
ctr = ctr + 1
Input #1, website(ctr)
Loop
ctr = 0
Do While Not EOF(2)
ctr = ctr + 1
Input #2, safety(ctr)
Loop
Picresults.Cls
For J = 1 To ctr
Picresults.Print safety(J); " is for "; website(J)
Close #1
Close #2
Next J
End Sub
