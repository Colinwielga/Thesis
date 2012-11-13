VERSION 5.00
Begin VB.Form frmMovies 
   Caption         =   "Form1"
   ClientHeight    =   11100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   Picture         =   "frmMovies.frx":0000
   ScaleHeight     =   11100
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   3360
      ScaleHeight     =   4995
      ScaleWidth      =   3915
      TabIndex        =   5
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort by Highest Gross"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdDirector 
      Caption         =   "Who Directed It"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdShowMovies 
      Caption         =   "Show Movies"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDirector_Click()
    Dim Picture(1 To 6) As String
    Dim Director(1 To 6) As String
    Dim Gross(1 To 6) As Single
    Dim Ctr, I As Integer
    Ctr = 0
    Open App.Path & "\movies.txt" For Input As #1
    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Picture(Ctr), Director(Ctr), Gross(Ctr)
    Loop
    Close #1
    picResults.Cls
    picResults.Print "Vince's Movies"; Tab(30); "Director"
    picResults.Print "*************************"
    For I = 1 To Ctr
        picResults.Print Picture(I); Tab(30); Director(I)
    Next I
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturn_Click()
    frmMovies.Hide
    frmMain.Show
End Sub

Private Sub cmdShowMovies_Click()
    Dim Picture(1 To 6) As String
    Dim Director(1 To 6) As String
    Dim Gross(1 To 6) As Single
    Dim Ctr, I As Integer
    Ctr = 0
    Open App.Path & "\movies.txt" For Input As #1
    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Picture(Ctr), Director(Ctr), Gross(Ctr)
    Loop
    Close #1
    picResults.Print "Vince's Movies"
    picResults.Print "**************"
    For I = 1 To Ctr
        picResults.Print Picture(I)
    Next I
   
End Sub

Private Sub cmdSort_Click()
    Dim Picture(1 To 6) As String
    Dim Director(1 To 6) As String
    Dim Gross(1 To 6) As Single
    Dim Ctr, I As Integer
    Ctr = 0
    Open App.Path & "\movies.txt" For Input As #1
    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Picture(Ctr), Director(Ctr), Gross(Ctr)
    Loop
    Close #1
    Dim pass, pos As Integer
    Dim tempMovie As String
    Dim tempPay As Single
    picResults.Print
    picResults.Print
    picResults.Print "Year"; Tab(30); "Box Office Gross"
    picResults.Print "*************************************************"
    For pass = 1 To Ctr - 1
        For pos = 1 To Ctr - pass
            If Gross(pos) < Gross(pos + 1) Then
                tempPay = Gross(pos)
                Gross(pos) = Gross(pos + 1)
                Gross(pos + 1) = tempPay
                tempMovie = Picture(pos)
                Picture(pos) = Picture(pos + 1)
                Picture(pos + 1) = tempMovie
            End If
        Next pos
    Next pass
    For I = 1 To Ctr
        picResults.Print Picture(I); Tab(30); FormatNumber(Gross(I), 0)
    Next I

End Sub
