VERSION 5.00
Begin VB.Form frmPlaces800 
   BackColor       =   &H00FF0000&
   Caption         =   """P""laces on the Map for $800"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8760
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNoBuzzPlaces800 
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton cmdBuzzPlaces800 
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
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label lblPlaces800 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "This country is the sixth most populous country in the world and is the second most populous country with a Muslim majority."
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
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmPlaces800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuzzPlaces800_Click()

    'Declaring local variables
    Dim QuestionPlaces800 As String
    Dim CorrectPlaces800 As String

    'Asking user for question and declaring correct question
    QuestionPlaces800 = InputBox("Enter your question", "Question to Places on the Map for $800")
    CorrectPlaces800 = CorrectQuestions(24)
    
    'Comparing user's question with correct question
    If LCase(QuestionPlaces800) = LCase(CorrectPlaces800) Then
        Winnings = Winnings + 800
        MsgBox "That is the correct question", , "Correct Question"
    Else
        Winnings = Winnings - 800
        MsgBox "That is incorrect. The correct question is" & " " & CorrectQuestions(24), , "Incorrect Question"
    End If
    
    'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hides forms
            frmPlaces800.Hide
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
            frmPlaces800.Hide
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

Private Sub cmdNoBuzzPlaces800_Click()

    'Lets the user know the correct answer and that he/she does not gain or lose any money
    MsgBox "The correct question is" & " " & CorrectQuestions(24) & "." & " " & "You do not gain any money or lose any money!!!", , "Correct Question"
    
        'Bring up the user's Character to show them new money total
    Select Case Player
        
        Case Is = 1
            
            'Shows and hides the forms
            frmPlaces800.Hide
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
            frmPlaces800.Hide
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
