VERSION 5.00
Begin VB.Form frmLevel4 
   BackColor       =   &H00C0C000&
   Caption         =   "Level 4"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuestion 
      Caption         =   "Get Question!"
      Height          =   1215
      Left            =   480
      TabIndex        =   6
      Top             =   5400
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFF00&
      Height          =   615
      Left            =   360
      ScaleHeight     =   555
      ScaleWidth      =   5235
      TabIndex        =   5
      Top             =   2160
      Width           =   5295
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Check Answer!"
      Height          =   1215
      Left            =   3240
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox txtBlue 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtRed 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   4485
      Left            =   5760
      Picture         =   "frmLevel4.frx":0000
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblDirections 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   $"frmLevel4.frx":715B
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   7815
   End
   Begin VB.Image imgQuit 
      Height          =   705
      Left            =   7080
      Picture         =   "frmLevel4.frx":7227
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label lblBlue 
      BackColor       =   &H00C0C000&
      Caption         =   "How Many Blue Fish?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblRed 
      BackColor       =   &H00C0C000&
      Caption         =   "How many Red Fish?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
End
Attribute VB_Name = "frmLevel4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NumArray(1 To 10) As Integer, Counter As Integer, QuestNum As Integer

Private Sub Form_Load()
    Dim numbers As Integer
    
    Open App.Path & "\Numbers.txt" For Input As #1
    Do Until EOF(1)
        Input #1, numbers
        Counter = Counter + 1
        NumArray(Counter) = numbers
    Loop
    Close #1
End Sub

Private Sub cmdQuestion_Click()
    
    If QuestNum < Counter Then
        picResults.Cls
        QuestNum = QuestNum + 1
        picResults.Print "How many of each fish do we need to get "; NumArray(QuestNum); " fish?"
    Else
        QuestNum = 0
            MsgBox "You've finished the Quiz!"
        frmLevel4.Visible = False
        frmLevel5.Visible = True
    End If
    
End Sub

Private Sub cmdEnter_Click()
    Dim Red As Integer, Blue As Integer, Total As Integer
    
    Red = txtRed.Text
    Blue = txtBlue.Text
    Total = Red + Blue
    
    Select Case Total
        Case Is = NumArray(QuestNum)
            MsgBox YourName & " You are correct!", , "Hooray!!"
        Case Else
            MsgBox YourName & " Try again!", , "Oops!"
    End Select
End Sub

Private Sub imgQuit_Click()
End
End Sub

