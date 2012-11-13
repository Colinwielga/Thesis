VERSION 5.00
Begin VB.Form frmQuestion4 
   BackColor       =   &H003D30AD&
   Caption         =   "Question 4"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   13125
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   3495
      Left            =   6600
      Picture         =   "frmQuestion4.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   5955
      TabIndex        =   7
      Top             =   960
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   4080
      Picture         =   "frmQuestion4.frx":4EF7
      ScaleHeight     =   3675
      ScaleWidth      =   4515
      TabIndex        =   6
      Top             =   4680
      Width           =   4575
   End
   Begin VB.OptionButton OptQ4 
      BackColor       =   &H003D30AD&
      Caption         =   "Performed at a Dinner Theater"
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
      Left            =   720
      TabIndex        =   5
      Top             =   3840
      Width           =   5175
   End
   Begin VB.OptionButton OptQ4 
      BackColor       =   &H003D30AD&
      Caption         =   "Having a Bad Hair Day"
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
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   5055
   End
   Begin VB.CommandButton cmdTo5 
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
      Left            =   720
      TabIndex        =   3
      Top             =   6480
      Width           =   2655
   End
   Begin VB.OptionButton OptQ4 
      BackColor       =   &H003D30AD&
      Caption         =   "Falling in Love"
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
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   4815
   End
   Begin VB.OptionButton OptQ4 
      BackColor       =   &H003D30AD&
      Caption         =   "Being Trapped in a Magical Castle"
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
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H003D30AD&
      Caption         =   "What's the Craziest thing that has ever happened to you?"
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
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   8775
   End
End
Attribute VB_Name = "frmQuestion4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTo5_Click()
'if statement to add 1 to the a specific counter depending on which option is selected

If OptQ4(0) = True Then
    CtrA = CtrA + 1
ElseIf OptQ4(1) = True Then
    CtrB = CtrB + 1
ElseIf OptQ4(2) = True Then
    CtrC = CtrC + 1
ElseIf OptQ4(3) = True Then
    CtrD = CtrD + 1
End If

'moves to next form
frmQuestion4.Hide
frmQuestion5.Show
End Sub

