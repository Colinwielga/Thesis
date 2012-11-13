VERSION 5.00
Begin VB.Form frmSources 
   Caption         =   "Sources"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   Picture         =   "frmSources.frx":0000
   ScaleHeight     =   8145
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCredit 
      Caption         =   "Sources"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   7440
      Width           =   1935
   End
   Begin VB.PictureBox picCredit 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   1320
      ScaleHeight     =   6675
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5880
      TabIndex        =   0
      Top             =   7440
      Width           =   1935
   End
End
Attribute VB_Name = "frmSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quest for The Cup~Minnesota Wild Trivia Game


'Author: Chris Adams

'Date: November 2007

'This form lists all Sources used in the creation of this project


Private Sub cmdCredit_Click()

'Declare variables
Dim Sources(1 To 100) As String
Dim SNum As Integer

cmdQuit.Enabled = True      'Quit button becomes functional
SNum = 0                    'Set initial value of SNum

    'Open the file that contains the sources for the project and load it into the array
    Open App.Path & "\files\credit.txt" For Input As #5
        Do Until EOF(5)
            SNum = SNum + 1
            Input #5, Sources(SNum)
            picCredit.Print Sources(SNum)
        Loop
    Close #5
    
End Sub

Private Sub cmdQuit_Click()

    'Quit the program
    End

End Sub
