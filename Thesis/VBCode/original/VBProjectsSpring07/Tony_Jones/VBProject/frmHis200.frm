VERSION 5.00
Begin VB.Form frmHis200 
   BackColor       =   &H00FF0000&
   Caption         =   "History for the Daily Double"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuzzHis200 
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblHis200 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "December 7th, 1941"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmHis200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuzzHis200_Click()
    
    'Declaring local variables
    Dim QuestionHis200 As String
    Dim CorrectHis200 As String
    
    'Asking user for question and declaring correct question
    QuestionHis200 = InputBox("Enter your question", "Question to History for the Daily Double")
    CorrectHis200 = CorrectQuestions(6)
    
    'Comparing user's question with correct question
    If LCase(QuestionHis200) = LCase(CorrectHis200) Then
        Winnings = Winnings + Wager
        MsgBox "That is the correct question", , "Correct Question"
    Else
        Winnings = Winnings - Wager
        MsgBox "That is incorrect. The correct question is" & " " & CorrectQuestions(6), , "Incorrect Question"
    End If
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hides forms
            frmHis200.Hide
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
            frmHis200.Hide
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
