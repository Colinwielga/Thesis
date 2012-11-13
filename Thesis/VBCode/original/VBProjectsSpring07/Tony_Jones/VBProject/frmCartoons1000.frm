VERSION 5.00
Begin VB.Form frmCartoons1000 
   BackColor       =   &H00FF0000&
   Caption         =   "Cartoon Characters for $1000"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNoBuzzCartoons1000 
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
      Height          =   1095
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdBuzzCartoons1000 
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
      Height          =   1095
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image imgCartoons1000 
      Height          =   3060
      Left            =   0
      Picture         =   "frmCartoons1000.frx":0000
      Top             =   0
      Width           =   3675
   End
End
Attribute VB_Name = "frmCartoons1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuzzCartoons1000_Click()

    'Declaring local variables
    Dim QuestionCartoons1000 As String
    Dim CorrectCartoons1000 As String
    
    'Asking user for question and declaring correct question
    QuestionCartoons1000 = InputBox("Enter your question", "Question to Cartoon Characters for $1000")
    CorrectCartoons1000 = CorrectQuestions(20)
    
    'Comparing user's question with correct question
    If LCase(QuestionCartoons1000) = LCase(CorrectCartoons1000) Then
        Winnings = Winnings + 1000
        MsgBox "That is the correct question", , "Correct Question"
    Else
        Winnings = Winnings - 1000
        MsgBox "That is incorrect. The correct question is" & " " & CorrectQuestions(20), , "Incorrect Question"
    End If
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and Hides the forms
            frmCartoons1000.Hide
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
        
            'Shows and Hides the forms
            frmCartoons1000.Hide
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

Private Sub cmdNoBuzzCartoons1000_Click()
    
    'Lets the user know the correct answer and that he/she does not gain or lose any money
    MsgBox "The correct question is" & " " & CorrectQuestions(20) & "." & " " & "You do not gain any money or lose any money!!!", , "Correct Question"
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hides forms
            frmCartoons1000.Hide
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
            
            'Shows and hides forms
            frmCartoons1000.Hide
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
