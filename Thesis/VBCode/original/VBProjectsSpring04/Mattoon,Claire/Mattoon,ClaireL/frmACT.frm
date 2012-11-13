VERSION 5.00
Begin VB.Form frmACT 
   BackColor       =   &H00FF8080&
   Caption         =   "Predicted GPA from ACT score"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FF80&
      Caption         =   "Quit"
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdRet 
      BackColor       =   &H0080C0FF&
      Caption         =   "Return to Welcome"
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnterACT 
      BackColor       =   &H008080FF&
      Caption         =   "Please click to enter your ACT score, and predict your college GPA"
      Height          =   1095
      Left            =   240
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Claire L. Mattoon"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M(1 To 3) As Single, SD(1 To 3) As Single, Ctr As Integer
Dim r As Single
Private Sub cmdEnterACT_Click()
frmACT.Cls
r = 0.6
Score = InputBox("Please type your ACT score and hit 'enter'")
GPA = ((((Score - M(2)) / SD(2)) * r) * SD(3)) + M(3)
Print "The predicted college GPA for a score of "; Score; " is "; FormatNumber(GPA, 2)
If GPA >= 3.5 Then
MsgBox "Congratulations! Based on your ACT score, you are likely to be on the Honor's Roll!", , "Congratulations!"
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
If Ctr = 0 Then
Path = "N:\CS130\handin\Mattoon, Claire\Mattoon, Claire L\"
Open Path & "Stats.txt" For Input As #1
End If
For Ctr = 1 To 3
    Input #1, M(Ctr), SD(Ctr)
Next Ctr
End Sub
