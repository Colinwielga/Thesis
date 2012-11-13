VERSION 5.00
Begin VB.Form frmQuestion2 
   BackColor       =   &H003D30AD&
   Caption         =   "Question 2"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   2775
      Left            =   8160
      Picture         =   "frmQuestion2.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   3555
      TabIndex        =   9
      Top             =   600
      Width           =   3615
   End
   Begin VB.PictureBox Picture3 
      Height          =   3015
      Left            =   4080
      Picture         =   "frmQuestion2.frx":6E80
      ScaleHeight     =   2955
      ScaleWidth      =   3555
      TabIndex        =   8
      Top             =   4440
      Width           =   3615
   End
   Begin VB.PictureBox Picture2 
      Height          =   4215
      Left            =   8160
      Picture         =   "frmQuestion2.frx":8B63
      ScaleHeight     =   4155
      ScaleWidth      =   3435
      TabIndex        =   7
      Top             =   3840
      Width           =   3495
   End
   Begin VB.OptionButton OptQ2 
      BackColor       =   &H003D30AD&
      Caption         =   "Toasted Marshmallows"
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
      Height          =   855
      Index           =   3
      Left            =   840
      TabIndex        =   6
      Top             =   3360
      Width           =   3495
   End
   Begin VB.OptionButton OptQ2 
      BackColor       =   &H003D30AD&
      Caption         =   "Whole Eggs"
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
      Left            =   840
      TabIndex        =   5
      Top             =   2760
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   4320
      Picture         =   "frmQuestion2.frx":EADD
      ScaleHeight     =   3075
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton cmdForm3 
      Caption         =   "Next>>"
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
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   2535
   End
   Begin VB.OptionButton OptQ2 
      BackColor       =   &H003D30AD&
      Caption         =   "A Light Salad"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.OptionButton OptQ2 
      BackColor       =   &H003D30AD&
      Caption         =   "Sloppy Joes"
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
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H003D30AD&
      Caption         =   "What is your favorite of the following foods?"
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
      Height          =   1215
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmQuestion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdForm3_Click()

'if statement to add 1 to the a specific counter depending on which option is selected

If OptQ2(0) = True Then
    CtrA = CtrA + 1
ElseIf OptQ2(1) = True Then
    CtrB = CtrB + 1
ElseIf OptQ2(2) = True Then
    CtrC = CtrC + 1
ElseIf OptQ2(3) = True Then
    CtrD = CtrD + 1
End If

'moves to next form
frmQuestion2.Hide
frmQuestion3.Show

End Sub
