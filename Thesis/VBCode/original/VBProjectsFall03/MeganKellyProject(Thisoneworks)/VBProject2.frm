VERSION 5.00
Begin VB.Form qualities23 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Mmmm... Salmon."
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox likeable 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Likeable"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox religious 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Religious"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox distracted 
      BackColor       =   &H00C0C0FF&
      Caption         =   "easily distracted by attractive people"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CheckBox scary 
      BackColor       =   &H00C0C0FF&
      Caption         =   "intimidating to others"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CheckBox sympoor 
      BackColor       =   &H00C0C0FF&
      Caption         =   "sympathetic to the less fortunate"
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CheckBox blame 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Tend to place blame on people."
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CheckBox group 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Are more concerned with the good of the country /group rather than the rights of individuals."
      Height          =   975
      Left            =   3240
      TabIndex        =   4
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CheckBox humanlife 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Believe that ALL human life is sacred or at least intrinsically valuable."
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CheckBox minions 
      BackColor       =   &H00C0C0FF&
      Caption         =   "You require minions."
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton Continue3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Continue"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton Quit3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Quit"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Please check all boxes that apply to you."
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "qualities23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Continue3_Click()
' Which not so nice person would you like to beat up today? "Megan'sVBProject.vbp"
'                       Intro1 (VBProject4.frm)
'                       Megan Kelly 11/03/03
' Purpose:  The purpose of this form is to collect information about the person and calculate it into a running sum to help determine the completely unscientific outcome of this exercise.

If likeable.Value = 1 Then sum = sum + 1
If religious.Value = 1 Then sum = sum + 1
If distracted.Value = 1 Then sum = sum - 3
If scary.Value = 1 Then sum = sum + 1
If sympoor.Value = 1 Then sum = sum + 1
If blame.Value = 1 Then sum = sum + 1
If humanlife.Value = 1 Then sum = sum + 1
If group.Value = 1 Then sum = sum + 1
If minions.Value = 1 Then sum = sum - 3
qualities23.Visible = False
qualities4.Visible = True
End Sub


Private Sub Quit3_Click()
End
End Sub
