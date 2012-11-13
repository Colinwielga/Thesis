VERSION 5.00
Begin VB.Form frmLit1000 
   BackColor       =   &H00FF0000&
   Caption         =   "Literature Authors for $1000"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNoBuzzLit1000 
      BackColor       =   &H0000FF00&
      Caption         =   "Don't Buzz In"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   3495
   End
   Begin VB.CommandButton cmdBuzzLit1000 
      BackColor       =   &H0000FF00&
      Caption         =   "Buzz In"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Image imgCanterburyTales 
      Height          =   4620
      Left            =   840
      Picture         =   "frmLit1000.frx":0000
      Top             =   0
      Width           =   7485
   End
End
Attribute VB_Name = "frmLit1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuzzLit1000_Click()
    
    'Declaring local variables
    Dim QuestionLit1000 As String
    Dim CorrectLit1000 As String

    'Asking user for question and declaring correct question
    QuestionLit1000 = InputBox("Enter your question", "Question to Literature Authors for $1000")
    CorrectLit1000 = CorrectQuestions(5)
    
    'Comparing user's question with correct question
    If LCase(QuestionLit1000) = LCase(CorrectLit1000) Then
        Winnings = Winnings + 1000
        MsgBox "That is the correct question", , "Correct Question"
    Else
        Winnings = Winnings - 1000
        MsgBox "That is incorrect. The correct question is" & " " & CorrectQuestions(5), , "Incorrect Question"
    End If
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hides forms
            frmLit1000.Hide
            frmKenMoney.Show
                        
            'Keeps the user's name
            frmKenMoney.picName.Cls
            frmKenMoney.picName.Print FName
            
            'Displays amount of winnings
            If Winnings >= 0 Then
                frmKenMoney.picWinnings.Cls
                frmKenMoney.picWinnings.Print FormatCurrency(Winnings, 0)
            Else
                frmKenMoney.picWinnings.Cls
                frmKenMoney.picWinnings.Print "-"; FormatCurrency(-Winnings, 0)
            End If
            
        Case Is = 2
            
            'Shows and hides the forms
            frmLit1000.Hide
            frmBushMoney.Show
            
            'Keeps the user's name
            frmBushMoney.picName.Cls
            frmBushMoney.picName.Print FName
            
            'Displays amount of winnings
            If Winnings >= 0 Then
                frmBushMoney.picWinnings.Cls
                frmBushMoney.picWinnings.Print FormatCurrency(Winnings, 0)
            Else
                frmBushMoney.picWinnings.Cls
                frmBushMoney.picWinnings.Print "-"; FormatCurrency(-Winnings, 0)
            End If
            
    End Select
    
End Sub

Private Sub cmdNoBuzzLit1000_Click()

    'Lets the user know the correct answer and that he/she does not gain or lose any money
    MsgBox "The correct question is" & " " & CorrectQuestions(5) & "." & " " & "You do not gain any money or lose any money!!!", , "Correct Question"
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hides the forms
            frmLit1000.Hide
            frmKenMoney.Show
            
            'Keeps the user's name
            frmKenMoney.picName.Cls
            frmKenMoney.picName.Print FName
            
            'Displays amount of winnings
            If Winnings >= 0 Then
                frmKenMoney.picWinnings.Cls
                frmKenMoney.picWinnings.Print FormatCurrency(Winnings, 0)
            Else
                frmKenMoney.picWinnings.Cls
                frmKenMoney.picWinnings.Print "-"; FormatCurrency(-Winnings, 0)
            End If
            
        Case Is = 2
            
            'Shows and hides the forms
            frmLit1000.Hide
            frmBushMoney.Show
            
            'Keeps the user's name
            frmBushMoney.picName.Cls
            frmBushMoney.picName.Print FName
            
            'Displays amount of winnings
            If Winnings >= 0 Then
                frmBushMoney.picWinnings.Cls
                frmBushMoney.picWinnings.Print FormatCurrency(Winnings, 0)
            Else
                frmBushMoney.picWinnings.Cls
                frmBushMoney.picWinnings.Print "-"; FormatCurrency(-Winnings, 0)
            End If
            
    End Select
    
End Sub
