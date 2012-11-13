VERSION 5.00
Begin VB.Form frmNCAA 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   6000
      TabIndex        =   12
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Selection"
      Height          =   855
      Left            =   3480
      TabIndex        =   11
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   855
      Left            =   1080
      TabIndex        =   10
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtInterceptions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   9
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtTouchdowns 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   8
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtYards 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   7
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtCompletions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtAttempts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblInterceptions 
      Caption         =   "Interceptions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblTouchdowns 
      Caption         =   "Touchdowns:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblYards 
      Caption         =   "Yards:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label lblCompletions 
      Caption         =   "Completions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblAttempts 
      Caption         =   "Pass Attempts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frmNCAA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()
'dim the neccessary variables
Dim Completions As Integer
Dim Attempts As Integer
Dim Touchdowns As Integer
Dim Interceptions As Integer
Dim Yards As Integer
Dim QBRating As Integer


'assign the user input values to the variables
Completions = txtCompletions
Attempts = txtAttempts
Touchdowns = txtTouchdowns
Interceptions = txtInterceptions
Yards = txtYards

'complete preliminary caculations

Yards = Yards * 8.4
Touchdowns = Touchdowns * 330
Interceptions = Interceptions * 200
Completions = Completions * 100

'finishing the equation

QBRating = Yards + Touchdowns - Interceptions + Completions
QBRating = QBRating / Attempts

'display the Rating in a message box

If QBRating > 200 Then
    MsgBox "Wow what a performance! Your QB's rating was" & (FormatNumber(QBRating, 1))
ElseIf QBRating > 150 Then
    MsgBox "A very great performance! Your QB's rating was" & (FormatNumber(QBRating, 1))
ElseIf QBRating > 120 Then
    MsgBox "An Average performance! Your QB's rating was" & (FormatNumber(QBRating, 1))
ElseIf QBRating <= 120 Then
    MsgBox "Its time to look at a backup! Your QB's rating was" & (FormatNumber(QBRating, 1))
End If




End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()

frmNFL.Hide
frmNCAA.Hide
frmStart.Show

End Sub

