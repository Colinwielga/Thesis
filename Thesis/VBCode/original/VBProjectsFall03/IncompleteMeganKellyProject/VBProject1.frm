VERSION 5.00
Begin VB.Form qualities4 
   BackColor       =   &H80000007&
   Caption         =   "Form4"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   LinkTopic       =   "Form4"
   ScaleHeight     =   5430
   ScaleWidth      =   5505
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Quit4 
      BackColor       =   &H00FF80FF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Continue4 
      BackColor       =   &H00FF80FF&
      Caption         =   "Continue"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CheckBox NRA 
      BackColor       =   &H80000007&
      Caption         =   "A member of the National Rifle Association"
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CheckBox wealthy 
      BackColor       =   &H80000007&
      Caption         =   "independently wealthy"
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CheckBox confident 
      BackColor       =   &H80000007&
      Caption         =   "self-confident"
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CheckBox athletic 
      BackColor       =   &H80000007&
      Caption         =   "athletic"
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CheckBox conservative 
      BackColor       =   &H80000007&
      Caption         =   "a political conservative (A Republican, or other right-wing affiliation)"
      ForeColor       =   &H00FF80FF&
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CheckBox strong 
      BackColor       =   &H80000007&
      Caption         =   "strong"
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CheckBox intelligent 
      BackColor       =   &H80000007&
      Caption         =   "intelligent"
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CheckBox healthy 
      BackColor       =   &H80000007&
      Caption         =   "healthy"
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Please check all qualities that describe you:"
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   4815
   End
End
Attribute VB_Name = "qualities4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Continue4_Click()
If healthy.Value = 1 Then sum = sum + 1
If intelligent.Value = 1 Then sum = sum + 1
If strong.Value = 1 Then sum = sum + 1
If conservative.Value = 1 Then sum = sum - 1
If athletic.Value = 1 Then sum = sum + 1
If confident.Value = 1 Then sum = sum + 1
If wealthy.Value = 1 Then sum = sum + 3
If NRA.Value = 1 Then sum = sum + 1
MsgBox (sum)
qualities4.Visible = False
vitals5.Visible = True
End Sub


Private Sub Quit4_Click()
End
End Sub
