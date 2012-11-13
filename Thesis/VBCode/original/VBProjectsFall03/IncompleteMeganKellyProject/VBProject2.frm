VERSION 5.00
Begin VB.Form qualities23 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox likeable 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Likeable"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox religious 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Religious"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.CheckBox distracted 
      BackColor       =   &H00C0C0FF&
      Caption         =   "easily distracted by attractive people"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CheckBox scary 
      BackColor       =   &H00C0C0FF&
      Caption         =   "intimidating to others"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CheckBox sympoor 
      BackColor       =   &H00C0C0FF&
      Caption         =   "sympathetic to the poor"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Tend to place blame on people."
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   0
      Width           =   2415
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Are more concerned with the good of the country /group rather than the rights of individuals."
      Height          =   975
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   3015
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Believe that all life is sacred or at least intrinsically valuable."
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00C0C0FF&
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton Continue3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Continue"
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Quit3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Quit"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "qualities23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Continue3_Click()

If likeable.Value = 1 Then sum = sum + 1
If religious.Value = 1 Then sum = sum + 1
If distracted.Value = 1 Then sum = sum + 1
If scary.Value = 1 Then sum = sum + 1
If sympoor.Value = 1 Then sum = sum + 1
If Check6.Value = 1 Then sum = sum + 1
If Check7.Value = 1 Then sum = sum + 1
If Check8.Value = 1 Then sum = sum + 1
If Check9.Value = 1 Then sum = sum + 1
qualities23.Visible = False
qualities4.Visible = True
End Sub


Private Sub Quit3_Click()
End
End Sub
