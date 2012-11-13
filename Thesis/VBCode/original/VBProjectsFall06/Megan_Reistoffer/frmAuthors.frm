VERSION 5.00
Begin VB.Form frmAuthors 
   BackColor       =   &H80000009&
   Caption         =   "Think you know something about authors?"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDickens 
      Caption         =   "Charles Dickens"
      Height          =   615
      Left            =   9360
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdFlies 
      Caption         =   "Lord of the Flies"
      Height          =   615
      Left            =   6480
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlath 
      Caption         =   "Sylvia Plath"
      Height          =   615
      Left            =   4200
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdDante 
      Caption         =   "Dante Alighieri"
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdPeter 
      BackColor       =   &H80000009&
      Caption         =   "Peter Rabbit"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2280
      Picture         =   "frmAuthors.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   6480
      Picture         =   "frmAuthors.frx":084B
      ScaleHeight     =   1635
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox picDickens 
      Height          =   1575
      Left            =   9360
      Picture         =   "frmAuthors.frx":1917
      ScaleHeight     =   1515
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox picPlath 
      Height          =   1455
      Left            =   4320
      Picture         =   "frmAuthors.frx":2223
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox picBPotter 
      Height          =   1815
      Left            =   240
      Picture         =   "frmAuthors.frx":2C31
      ScaleHeight     =   1755
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdTrivias 
      Caption         =   "Try out some Trivia?"
      Height          =   855
      Left            =   4560
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdBeginning 
      Caption         =   "Return to the Beginning"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "My Head hurts- I quit."
      Height          =   855
      Left            =   7920
      Picture         =   "frmAuthors.frx":374B
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
End
Attribute VB_Name = "frmAuthors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBeginning_Click()
    'go to intro form
    frmWhat.Visible = True
    frmAuthors.Visible = False
    
End Sub

Private Sub cmdDante_Click()
Dim A As String
'ask user question via inputbox and disply answer via message box
A = InputBox("What work of Dante's charts a journey through Hell, Purgatory and Paradise?", "Question")
If A = "The Divine Comedy" Then
    MsgBox "Way to Go!", , "Bravo!"
Else
    MsgBox "Incorrect, try again!", , "Wrong!"
End If
End Sub

Private Sub cmdDickens_Click()
Dim A As String
'ask user question via inputbox and disply answer via message box
A = InputBox("Which Dickens novel features a main character named Pip?", "Question")
If A = "Great Expectations" Then
    MsgBox "Way to Go!", , "Bravo!"
Else
    MsgBox "Incorrect, try again!", , "Wrong!"
End If
End Sub

Private Sub cmdFlies_Click()
Dim A As String
'ask user question via inputbox and disply answer via message box
A = InputBox("Who wrote Lord of the Flies?", "Question")
If A = "William Golding" Then
    MsgBox "Way to Go!", , "Bravo!"
Else
    MsgBox "Incorrect, try again!", , "Wrong!"
End If
End Sub

Private Sub cmdPeter_Click()
Dim A As String
'ask user question via inputbox and disply answer via message box
A = InputBox("Who wrote Peter Rabbit?", "Question")
If A = "Beatrix Potter" Then
    MsgBox "Way to Go!", , "Bravo!"
Else
    MsgBox "Incorrect, try again!", , "Wrong!"
End If

End Sub


Private Sub cmdPlath_Click()
Dim A As String
'ask user question via inputbox and disply answer via message box
A = InputBox("What semi-autobiographical novel did Plath write about her suicide attempts?", "Question")
If A = "The Bell Jar" Then
    MsgBox "Way to Go!", , "Bravo!"
Else
    MsgBox "Incorrect, try again!", , "Wrong!"
End If
End Sub

Private Sub cmdQuit_Click()
    End
End Sub


Private Sub cmdTrivias_Click()
    'go to trivia form
    frmTrivia.Visible = True
    frmAuthors.Visible = False
    
End Sub

