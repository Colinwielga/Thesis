VERSION 5.00
Begin VB.Form frmSAT 
   BackColor       =   &H008080FF&
   Caption         =   "Predicted GPA from SAT score"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdRet 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to Welcome"
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnterSAT 
      BackColor       =   &H0080FF80&
      Caption         =   "Please click to enter your SAT score, and compute your predicted college GPA"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "Claire L. Mattoon"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "frmSAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M(1 To 3) As Single, SD(1 To 3) As Single, Ctr As Integer
Dim r As Single

Private Sub cmdEnterSAT_Click()
frmSAT.Cls
r = 0.4
Score = InputBox("Please type your SAT score and hit 'enter'")
GPA = ((((Score - M(1)) / SD(1)) * r) * SD(3)) + M(3)
Print "The predicted college GPA for a score of "; Score; " is "; FormatNumber(GPA, 2)
If GPA >= 3.5 Then
MsgBox "Congratulations! Based on your SAT score, you are likely to be on the Honor's Roll!", , "Congratulations!"
End If
cmdQuit.Visible = True
cmdQuit.Enabled = True
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRet_Click()
frmwelcome.Show
End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\Mattoon, Claire\Mattoon, Claire L\"
Open Path & "Stats.txt" For Input As #2
Ctr = 0
For Ctr = 1 To 3
    Input #2, M(Ctr), SD(Ctr)
Next Ctr
End Sub
