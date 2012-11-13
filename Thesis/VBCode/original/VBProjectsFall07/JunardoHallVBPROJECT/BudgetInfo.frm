VERSION 5.00
Begin VB.Form BudgetInfo 
   BackColor       =   &H00000000&
   Caption         =   "BudgetInfo"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000000FF&
      Caption         =   "Start Over"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdCompare 
      BackColor       =   &H000000FF&
      Caption         =   "Has Budget Requirement Been Met?"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H000000FF&
      Caption         =   "Calculate Total Costs/Expenses"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdCost 
      BackColor       =   &H000000FF&
      Caption         =   "Enter Estimated Costs Per Month  (9 Months)"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H000000FF&
      Height          =   6255
      Left            =   360
      ScaleHeight     =   6195
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   1560
      Width           =   5535
   End
   Begin VB.TextBox txtBudget 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblGeneral 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Budget Information For '07-'08 School Year"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label lblBudget 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Enter Budget Amount:"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "BudgetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'declare variables
Dim COct, CNov, Dec, CJan, CFeb, CMarch, CApril, CMay As Single
Dim sum As Single
Dim Budget As Single
Dim ctr, e As Integer
Dim Expense(1 To 9) As String




Private Sub cmdCost_Click()
    'this button will cause a series of input boxes to appear,
    'these boxes will ask the user to enter the estimated costs of expenses for
    'the various months within the school year
    

'ask for user input
e = 0
For e = 1 To 9
    Expense(e) = InputBox("Enter the cost of expenses per month", "Expenses")
    pbxResults.Print "Month"; e; ":", Expense(e)
Next e


End Sub
Private Sub cmdTotal_Click()
    'this button will calculate the sum of the costs entered by the user

cmdCost.Visible = False 'this will cause the button to become inactive

Budget = txtBudget.Text

'make calculation
sum = 0
For e = 1 To 9
    sum = sum + Expense(e)
Next e


'print results
pbxResults.Print "**********************************************************************"
pbxResults.Print "The budget amount is "; FormatCurrency(Budget, 2)
pbxResults.Print "The total of estimated costs is "; FormatCurrency(sum, 2)


End Sub
Private Sub cmdCompare_Click()
    'this button will compare the sum of expense with the budget amount entered by the user
    ' the result will determine whether the user has met their budget requirement or not

cmdTotal.Visible = False    'this will cause the button to become inactive
cmdCompare.Visible = False


'declare variables
Dim remainder As Single
Dim over As Single

pbxResults.Print "*************************************************************"

'calculation to determine if the user has met their budget requirement
remainder = 0
If sum < Budget Then
    remainder = Budget - sum
    pbxResults.Print "You have "; FormatCurrency(remainder, 2); " left in your budget."
    pbxResults.Print "This does not meet your budget requirement."
        ElseIf sum = Budget Then
        pbxResults.Print "You have met your budget requirement."
            ElseIf sum > Budget Then
            MsgBox "Your expenses exceed your budget.", , "Error"
End If

'inform user of how much they have exceeded their budget limit by
Select Case sum
    Case Is > Budget
    over = sum - Budget
    pbxResults.Print "You need to take out "; FormatCurrency(over, 2); " from expenses to meet budget requirement."
End Select

End Sub
Private Sub cmdClear_Click()
    'this button will allow the user to restart the program if the have found an error by:
    'clearing the information in the picture box
    'and causing the estimated cost button to reappear

pbxResults.Cls
cmdCost.Visible = True
cmdTotal.Visible = True
cmdCompare.Visible = True

End Sub


Private Sub cmdQuit_Click()
    'this buttion will end the program
End
End Sub

