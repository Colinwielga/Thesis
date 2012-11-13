VERSION 5.00
Begin VB.Form frmMath800 
   BackColor       =   &H00FF0000&
   Caption         =   "Derivatives for $800"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNoBuzzMath800 
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
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton cmdBuzzMath800 
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label lblMath800 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "e^x"
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
Attribute VB_Name = "frmMath800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuzzMath800_Click()

    'Declaring local variables
    Dim QuestionMath800 As String
    Dim CorrectMath800 As String

    'Asking user for question and declaring correct question
    QuestionMath800 = InputBox("Enter your question", "Question to Derivatives for $800")
    CorrectMath800 = CorrectQuestions(14)
    
    'Comparing user's question with correct question
    If LCase(QuestionMath800) = LCase(CorrectMath800) Then
        Winnings = Winnings + 800
        MsgBox "That is the correct question", , "Correct Question"
    Else
        Winnings = Winnings - 800
        MsgBox "That is incorrect. The correct question is" & " " & CorrectQuestions(14), , "Incorrect Question"
    End If
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hides forms
            frmMath800.Hide
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
            frmMath800.Hide
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

Private Sub cmdNoBuzzMath800_Click()

    'Lets the user know the correct answer and that he/she does not gain or lose any money
    MsgBox "The correct question is" & " " & CorrectQuestions(14) & "." & " " & "You do not gain any money or lose any money!!!", , "Correct Question"
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hides the forms
            frmMath800.Hide
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
            frmMath800.Hide
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
