VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form3"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form3"
   ScaleHeight     =   4980
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1215
      Left            =   6960
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Menu"
      Height          =   1335
      Left            =   6960
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdSources 
      Caption         =   "Show Sources"
      Height          =   1335
      Left            =   6960
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox PicBib 
      BackColor       =   &H80000009&
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3195
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   1320
      Width           =   6255
   End
   Begin VB.Label lblLibrary 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bibliography"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form shows the bibliography

Private Sub cmdMenu_Click()
'Back to Menu

Form3.Hide
Form1.Show
End Sub

Private Sub cmdQuit_Click()
'Quit

End
End Sub

Private Sub cmdSources_Click()
'Displays bibliography

Dim Text(1 To 25) As String, K As Integer

'Open data file
Open App.Path & "\Bibliography.txt" For Input As #2

'Set counter K to zero
K = 0

'Clear picture box
PicBib.Cls

'Read and print array
Do Until EOF(2)
    K = K + 1
    Input #2, Text(K)
    PicBib.Print Text(K)
Loop

Close #2

End Sub
