VERSION 5.00
Begin VB.Form StateMoney 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Hide 
      Caption         =   "Go to Inflation Calculator"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      TabIndex        =   6
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton Dollar 
      Caption         =   "Click to put incomes in decendnig order"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   480
      TabIndex        =   5
      Top             =   6120
      Width           =   2775
   End
   Begin VB.CommandButton Alpha 
      Caption         =   "Click to put all states into alphbetical order"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton PickState 
      Caption         =   "Choose State to find income"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000000&
      Height          =   11055
      Left            =   3600
      ScaleHeight     =   10995
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   9960
      Picture         =   "PickState.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   2280
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "MEDIAN HOUSEHOLD INCOMES IN           2002 AND AN INFLATION                                 CALCULATOR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   9600
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "StateMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Economics (M:\CS130\MatthewGoldade\VB_Project\Economics.vbp)
'Form 1 (StateMoney)
'Matthew Goldade
'Oct. 28, 2003
'The purpose of this program is to input a state and find the median household income for that particular state.
'It also uses the record of Consumer Price Index to find the purchasing power of a current dollar in a year chosen by the user.
Option Explicit
Public PATH As String
Dim State(1 To 51) As String, Income(1 To 51) As Double, F As Integer


Private Sub Alpha_Click()
Results.Cls
Dim Pass As Single, temp As String

Open PATH & "2001Incomes.txt" For Input As #1

    For F = 1 To 51
        Input #1, State(F), Income(F)
    Next F
    
    For Pass = 1 To 50
        For F = 1 To 51 - Pass
            If State(F) > State(F + 1) Then
                temp = State(F)
                State(F) = State(F + 1)
                State(F + 1) = temp
                temp = Income(F)
                Income(F) = Income(F + 1)
                Income(F + 1) = temp
            End If
        Next F
    Next Pass
    
Results.Print "State", "Median Incomes"
Results.Print "---------------------------------------------------------------"

    For F = 1 To 51
        Results.Print State(F), , FormatCurrency(Income(F), 0)
    Next F

Close #1

End Sub

Private Sub Dollar_Click()
Results.Cls
Dim Pass As Single, temp As String

Open PATH & "2001Incomes.txt" For Input As #1

    For F = 1 To 51
        Input #1, State(F), Income(F)
    Next F
    
    For Pass = 1 To 50
        For F = 1 To 51 - Pass
            If Income(F) < Income(F + 1) Then
                temp = Income(F)
                Income(F) = Income(F + 1)
                Income(F + 1) = temp
                temp = State(F)
                State(F) = State(F + 1)
                State(F + 1) = temp
            End If
        Next F
    Next Pass
    
Results.Print "State", "Median Incomes"
Results.Print "---------------------------------------------------------------"

    For F = 1 To 51
        Results.Print State(F), , FormatCurrency(Income(F), 0)
    Next F

Close #1

End Sub

Private Sub Form_Load()
PATH = "N:\CS130\handin\MatthewGoldade\"
End Sub

Private Sub Hide_Click()
Calculator.Show
StateMoney.Hide
End Sub

Private Sub PickState_Click()
Dim A As String, F As Integer, NotFound As Boolean, X As Integer

Results.Cls

Open PATH & "2001Incomes2.txt" For Input As #1
    For F = 1 To 51
        Input #1, State(F), Income(F)
    Next F
    
A = InputBox("Type in any state to find the median annual household income", , "Enter State")

F = 1
NotFound = True
    Do While NotFound
        If F >= 51 Then
            Exit Do
        Else
            If A = State(F) Then
            NotFound = False
            X = F
            Exit Do
            End If
            F = F + 1
        End If
        Loop
    If NotFound Then
        MsgBox "Sorry, but your entered a nonexistant state", , "Invalid State"
        MsgBox "Make sure you capitalize each state, and you must spell it correctly", , "Don't Forget"
    Else
        Results.Print "State", "Median Income"
        Results.Print "----------------------------------------------------"
        Results.Print State(F), Income(F)
    End If
    
Close #1
        
End Sub

