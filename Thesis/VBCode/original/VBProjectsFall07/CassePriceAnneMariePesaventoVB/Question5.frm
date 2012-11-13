VERSION 5.00
Begin VB.Form frmQuestion5 
   BackColor       =   &H003D30AD&
   Caption         =   "Question 5"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   14595
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   2295
      Left            =   9720
      Picture         =   "Question5.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   4515
      TabIndex        =   9
      Top             =   5400
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   8640
      Picture         =   "Question5.frx":267FA
      ScaleHeight     =   4035
      ScaleWidth      =   5355
      TabIndex        =   8
      Top             =   1080
      Width           =   5415
   End
   Begin VB.PictureBox Picture3 
      Height          =   3855
      Left            =   3120
      Picture         =   "Question5.frx":2C30A
      ScaleHeight     =   3795
      ScaleWidth      =   5235
      TabIndex        =   7
      Top             =   1080
      Width           =   5295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   5160
      Picture         =   "Question5.frx":31718
      ScaleHeight     =   2475
      ScaleWidth      =   4395
      TabIndex        =   6
      Top             =   5280
      Width           =   4455
   End
   Begin VB.OptionButton OptQ5 
      BackColor       =   &H003D30AD&
      Caption         =   "Passion"
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
      Width           =   2175
   End
   Begin VB.OptionButton OptQ5 
      BackColor       =   &H003D30AD&
      Caption         =   "Extraordinary Good-Looks"
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
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdNext 
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
      Height          =   1455
      Left            =   1680
      TabIndex        =   3
      Top             =   5280
      Width           =   3255
   End
   Begin VB.OptionButton OptQ5 
      BackColor       =   &H003D30AD&
      Caption         =   "Fangs"
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
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.OptionButton OptQ5 
      BackColor       =   &H003D30AD&
      Caption         =   "Kindness"
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
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H003D30AD&
      Caption         =   "What one thing makes you stand out from everybody else?"
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
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   8415
   End
End
Attribute VB_Name = "frmQuestion5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNext_Click()
If OptQ5(0) = True Then
    CtrA = CtrA + 1
ElseIf OptQ5(1) = True Then
    CtrB = CtrB + 1
ElseIf OptQ5(2) = True Then
    CtrC = CtrC + 1
ElseIf OptQ5(3) = True Then
    CtrD = CtrD + 1
End If


frmQuestion5.Hide
frmResults.Show

End Sub
