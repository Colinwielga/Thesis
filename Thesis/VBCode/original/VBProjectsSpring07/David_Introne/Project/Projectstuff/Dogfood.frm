VERSION 5.00
Begin VB.Form Dogfood 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Dog Food Choice"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   3600
      Left            =   5040
      Picture         =   "Dogfood.frx":0000
      Top             =   4560
      Width           =   3600
   End
   Begin VB.Image small 
      Height          =   3600
      Left            =   5160
      Picture         =   "Dogfood.frx":28E3
      Top             =   840
      Width           =   3600
   End
   Begin VB.Image eating 
      Height          =   3090
      Left            =   8760
      Picture         =   "Dogfood.frx":5086
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   4050
   End
   Begin VB.Image Large 
      Height          =   4455
      Left            =   9000
      Picture         =   "Dogfood.frx":962D
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   8040
      Left            =   0
      Picture         =   "Dogfood.frx":22991
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4905
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on a dog food to make your selcetion. "
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Dogfood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Image2_Click()
    If puppick = 12 Then 'if you have a particulare dog you must pick a particulare dog food
        MsgBox "Mmm, smells great, this is the right dog food for your puppy.", , " Result"
        Score = Score + 5
        MsgBox "You gain another 5 points", , Score
    Else
        MsgBox "Check your breed, this breed is either too big or it resembles a hotdog.", , " Result"
        Score = Score - 3
        MsgBox "You lost 3 points for feeding " & Pupname & " incorectly.", , "Score"
    End If
        Select Case puppick 'this automatically takes you back to the cooresponding dog profile
                Case 11
                    ProShep.Show
                    Dogfood.Hide
                Case 12
                    ProPit.Show
                    Dogfood.Hide
                Case 13
                    ProMtn.Show
                    Dogfood.Hide
                Case 14
                    Produch.Show
                    Dogfood.Hide
            End Select

End Sub

Private Sub Large_Click()
    If puppick = 11 Or puppick = 13 Then 'if you have a particulare dog you must pick a particulare dog food
        MsgBox "Mmm, smells great, this is the right dog food for your puppy", , " Result"
        Score = Score + 5
        MsgBox "You gain another 5 points", , Score
    Else
        MsgBox "Check your breed, this breed is too small.", , " Result"
        Score = Score - 3
        MsgBox "You lost 3 points for feeding " & Pupname & " incorectly.", , "Score"
    End If
        Select Case puppick
            Case 11 'this automatically takes you back to the cooresponding dog profile
                ProShep.Show
                Dogfood.Hide
            Case 12
                ProPit.Show
                Dogfood.Hide
            Case 13
                ProMtn.Show
                Dogfood.Hide
            Case 14
                Produch.Show
                Dogfood.Hide
        End Select

End Sub

Private Sub small_Click()
    If puppick = 14 Then 'if you have a particulare dog you must pick a particulare dog food
        MsgBox "Mmm, smells great, this is the right dog food for your puppy", , " Result"
        Score = Score + 5
        MsgBox "You gain another 5 points", , Score
    Else
        MsgBox "Check your breed, this breed too large.", , " Result"
        Score = Score - 3
        MsgBox "You lost 3 points for feeding " & Pupname & " incorectly.", , "Score"
    End If
        Select Case puppick
                Case 11 'this automatically takes you back to the cooresponding dog profile
                    ProShep.Show
                    Dogfood.Hide
                Case 12
                    ProPit.Show
                    Dogfood.Hide
                Case 13
                    ProMtn.Show
                    Dogfood.Hide
                Case 14
                    Produch.Show
                    Dogfood.Hide
            End Select
End Sub
