VERSION 5.00
Begin VB.Form frmQuestion3 
   BackColor       =   &H003D30AD&
   Caption         =   "Question 3"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptQ3 
      BackColor       =   &H003D30AD&
      Caption         =   "Scream && Run Away"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   735
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   4200
      Width           =   3855
   End
   Begin VB.OptionButton OptQ3 
      BackColor       =   &H003D30AD&
      Caption         =   "Shoot it"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   615
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   4920
      Picture         =   "frmQuestion3.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton cmdTo4 
      BackColor       =   &H8000000A&
      Caption         =   "Next>>>"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   2655
   End
   Begin VB.OptionButton OptQ3 
      BackColor       =   &H003D30AD&
      Caption         =   "Gently feed them bird seed and stroke their feathers"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   4575
   End
   Begin VB.OptionButton OptQ3 
      BackColor       =   &H003D30AD&
      Caption         =   "Growl Ferociously!"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H003D30AD&
      Caption         =   "What would you do if a bird                    landed on you?"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   1095
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmQuestion3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTo4_Click()

'if statement to add 1 to the a specific counter depending on which option is selected

If OptQ3(0) = True Then
    CtrA = CtrA + 1
ElseIf OptQ3(1) = True Then
    CtrB = CtrB + 1
ElseIf OptQ3(2) = True Then
    CtrC = CtrC + 1
ElseIf OptQ3(3) = True Then
    CtrD = CtrD + 1
End If

'moves to next form
frmQuestion3.Hide
frmQuestion4.Show

End Sub
