VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00008000&
   Caption         =   "InflationCalc"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form2"
   ScaleHeight     =   6165
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back to Incomes"
      Height          =   855
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox Results 
      Height          =   615
      Left            =   480
      ScaleHeight     =   555
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton InflationCalc 
      Caption         =   "Input Year for Dollar Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   " $   Inflation Calculator    $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Economics (M:\CS130\MatthewGoldade\VB_Project\Economics.vbp)
'Form 1 (StateMoney)
'Mattehw Goldade
'Oct. 28, 2003
'The purpose of this program is to input a state and find the median household income for that particular state.
'It also uses the record of Consumer Price Index to find the purchasing power of a current dollar in a year chosen by the user.
Option Explicit

Private Sub Command1_Click()
'Clear the picture box
Results.Cls

End Sub

Private Sub Command2_Click()
'Hide the current form and show the inflation calculator
StateMoney.Show
Calculator.Hide
End Sub

Private Sub InflationCalc_Click()
'Declaring all variables
Dim I As Double, A As Double, Year(1 To 90) As Double, CPI(1 To 90) As Single, Inflation As Single
Dim NotFound As Boolean, X As Integer, n As String
'Open fie and then read them into arrays
Open StateMoney.PATH & "YearlyCPI.txt" For Input As #1

    For I = 1 To 90
        Input #1, Year(I), CPI(I)
    Next I
'With input box, select a year to compute purchasing power
A = InputBox("Select a year to discover purchasing power.                 Example: 1978", "Choose a Year")

    I = 1
'Read input and find it in the array, if found print the dollar value, otherwise message box appears with instructions
NotFound = True
Do While NotFound
    If I >= 91 Then
       Exit Do
    Else
        If A = Year(I) Then
        NotFound = False
        X = I
        Exit Do
        End If
        I = I + 1
    End If
    Loop

If NotFound Then

        MsgBox "Sorry, but this year does not have a record of CPI", , "Invalid Year"
        MsgBox "Remember, that CPI was not calculated until the year 1913", , "Remember"

        Close #1

    Else
        Results.Print FormatCurrency(CPI(I) / 179.9); " in"; A; " is equal to $1 right now!!"
        Results.Print ""
    End If
'close file that information was read from
Close #1

End Sub

