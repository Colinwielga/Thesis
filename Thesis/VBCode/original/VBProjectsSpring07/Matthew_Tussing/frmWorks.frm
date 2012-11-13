VERSION 5.00
Begin VB.Form frmWorks 
   BackColor       =   &H000080FF&
   Caption         =   "Works Cited"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear the Pictrue Box (Works Cited)"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdGoback 
      Caption         =   "Go Back To Main Page"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdWorks 
      Caption         =   "See the Works Cited"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   5775
      Left            =   2520
      ScaleHeight     =   5715
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Label lblWorks 
      BackColor       =   &H00FFFF00&
      Caption         =   "         WORKS CITED FOR MATT'S VB PROGRAM  3/30/07"
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmWorks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    picResults.Cls 'clears the works cited
End Sub

Private Sub cmdGoback_Click()
    frmProject.Show 'goes back to the main screen
    frmWorks.Hide
End Sub

Private Sub cmdQuit_Click()
End 'ends the program
End Sub

Private Sub cmdWorks_Click()
Dim words(1 To 100) As String
Dim ctr As Integer
Dim A As Integer

ctr = 0

Open App.Path & "\works.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, words(ctr)
    Loop
    Close #1
    
    For A = 1 To ctr
        picResults.Print words(A)
    Next A

End Sub
