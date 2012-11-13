VERSION 5.00
Begin VB.Form FrmCompare 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Click to see companies that have the same Net Income as Phototec Inc."
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Click here to sort the competitors from highest Net Income to lowest"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Click here to load and print the data of Phototec Inc.'s competitors"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FFFF&
      Height          =   5655
      Left            =   1800
      ScaleHeight     =   5595
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   4080
      Width           =   855
   End
End
Attribute VB_Name = "FrmCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable vs. Absorption Cost Accounting(VBProject.vbp)
'FrmCompare(frmCompare.frm)
'Mike Rakes, 3/12
'The purpose of this form is to load a list of competitors of the ficitonal company Phototec.
'Then it will sort the list of competitors by total net income from highest to lowest, and last
'it willl ask the user to input their higher income from the previous form and will see if any of the companies
'have the same total net income as Phototec.
Option Explicit
Dim PATH As String, CTR As Integer, Company(1 To 5) As String
Dim Income(1 To 5) As Single, YourIncome As Single
Private Sub cmdLoad_Click()
'Prepares data to be loaded
PATH = "N:\CS130\handin\Rakes, Michael\"
CTR = 0
'Opens data and stores it in two arrays
Open PATH & "Income.txt" For Input As #1
'prints the data in a picture box
picResults.Print "Company", "Net Income"
picResults.Print "*********************************************"
    Do While Not EOF(1)
      CTR = CTR + 1
      Input #1, Company(CTR), Income(CTR)
      picResults.Print Company(CTR), FormatCurrency(Income(CTR))
    Loop
'makes GUI more user friendly
cmdLoad.Enabled = False
cmdSort.Enabled = True

End Sub

Private Sub cmdQuit_Click()
'Quits out of program
End
End Sub

Private Sub cmdSearch_Click()
'Dims variables and Gets your income from an input box
Dim j As Integer
Dim found As Boolean
YourIncome = InputBox("Enter your higher Net Income from the Income Statements")
found = False
'prints the companies that have the same net income as yours
picResults.Print
picResults.Print
picResults.Print "The companies that have the same Net Income as Phototec Inc., if any, are: "
picResults.Print "***************************************************************************"

For j = 1 To CTR
    If Income(j) = YourIncome Then
        picResults.Print
        picResults.Print Company(j)
        found = True
    End If
Next j

If Not found Then
    MsgBox "Sorry, no companies have the same Net Income as Phototec Inc.", , "Bummer!"
End If
'Makes GUI more user friendly
cmdSearch.Enabled = False

End Sub

Private Sub cmdSort_Click()
Dim tempCompany As String
Dim tempIncome As Single
Dim PASS As Integer, COMP As Integer, j As Integer


'use Bubble Sort to put incmoes in order from highest to lowest
For PASS = 1 To CTR - 1
    For COMP = 1 To CTR - PASS
        If Income(COMP) < Income(COMP + 1) Then
        
            'switch company names
            tempCompany = Company(COMP)
            Company(COMP) = Company(COMP + 1)
            Company(COMP + 1) = tempCompany
            
            'and also switch Incomes
            tempIncome = Income(COMP)
            Income(COMP) = Income(COMP + 1)
            Income(COMP + 1) = tempIncome
            
        End If
    Next COMP
Next PASS

picResults.Print
picResults.Print
picResults.Print "*******Income (Hightest to Lowest)**********"

'Display Companies in Numeric order (based on income) along with their Incomes
For j = 1 To CTR
    picResults.Print Company(j), FormatCurrency(Income(j))
Next j
'Makes GUI more user friendly
cmdSort.Enabled = False
cmdSearch.Enabled = True
End Sub

Private Sub Form_Load()
'Makes GUI more user friendly
cmdSort.Enabled = False
cmdSearch.Enabled = False
End Sub
