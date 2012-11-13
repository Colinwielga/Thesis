VERSION 5.00
Begin VB.Form frmPlaces600 
   BackColor       =   &H00FF0000&
   Caption         =   """P""laces on the Map for $600"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNoBuzzPlaces600 
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton cmdBuzzPlaces600 
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Image imgPlaces600 
      Height          =   4530
      Left            =   0
      Picture         =   "frmPlaces600.frx":0000
      Top             =   0
      Width           =   6585
   End
End
Attribute VB_Name = "frmPlaces600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuzzPlaces600_Click()

    'Declaring local variables
    Dim QuestionPlaces600 As String
    Dim CorrectPlaces600 As String

    'Asking user for question and declaring correct question
    QuestionPlaces600 = InputBox("Enter your question", "Question to Places on the Map for $600")
    CorrectPlaces600 = CorrectQuestions(23)
    
    'Comparing user's question with correct question
    If LCase(QuestionPlaces600) = LCase(CorrectPlaces600) Then
        Winnings = Winnings + 600
        MsgBox "That is the correct question", , "Correct Question"
    Else
        Winnings = Winnings - 600
        MsgBox "That is incorrect. The correct question is" & " " & CorrectQuestions(23), , "Incorrect Question"
    End If
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hdies forms
            frmPlaces600.Hide
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
            frmPlaces600.Hide
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

Private Sub cmdNoBuzzPlaces600_Click()

    'Lets the user know the correct answer and that he/she does not gain or lose any money
    MsgBox "The correct question is" & " " & CorrectQuestions(23) & "." & " " & "You do not gain any money or lose any money!!!", , "Correct Question"
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hides the forms
            frmPlaces600.Hide
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
            
            'Shows an hides the forms
            frmPlaces600.Hide
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
