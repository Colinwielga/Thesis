VERSION 5.00
Begin VB.Form frmSorter 
   BackColor       =   &H00000000&
   Caption         =   "Currency Sorter"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   3255
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000C000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   3255
   End
   Begin VB.CommandButton cmdValue 
      BackColor       =   &H0000C000&
      Caption         =   "Sort by Monetary Value"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H0000C000&
      Caption         =   "Sort by Name"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H0000C000&
      Caption         =   "Load Currencies"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000B&
      Height          =   5175
      Left            =   2520
      ScaleHeight     =   5115
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   2400
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "* All monetary values are displayed in regards to an American Dollar. for example: 1 British Pound = 1.96 American Dollars."
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1185
      Left            =   240
      Picture         =   "frmSorter.frx":0000
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   2100
   End
   Begin VB.Label lblSorter 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   $"frmSorter.frx":11723
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Here the user can load and sort multiple world currencies.  Each currency has a moneatry value
'relative to an american dollar.  the rates are up-to-date and accurate.  First the user loads
'the data from a file, and then they can either sort the data in alphabetical order by name or
'Ascending order by monetary value.

Dim CurrenciesNames(1 To 100) As String     'Sets variables
Dim CurrenciesValue(1 To 100) As Single
Dim Values As Single
Dim Currencies As String
Dim ctr As Integer

Private Sub cmdLoad_Click()
    
    picResults.Print "********************************World Currencies******************************************"
Open App.Path & "\Currencies.txt" For Input As #1       'Opens file as input

Do While Not EOF(1)
    ctr = ctr + 1       'Adds one to the counter
    Input #1, CurrenciesNames(ctr), CurrenciesValue(ctr)        'Inputs two seperate arrays
        picResults.Print CurrenciesNames(ctr), , FormatCurrency(CurrenciesValue(ctr), 2)        'prints the arrays as loaded
Loop        'Loops back to the Do while not the end of file
Close #1        'Closes the file
End Sub

Private Sub cmdName_Click()
Dim pass As Integer     'Sets variables
Dim pos As Integer
Dim tempName As String
Dim tempValue As Single
Dim X As Integer

pos = 0     'Initiates varables
pass = 0
tempName = 0
tempValue = 0
X = 0
ctr = 0
  
  picResults.Print "********************************World Currencies******************************************"
Open App.Path & "\Currencies.txt" For Input As #1       'Opens file as input as loaded

Do While Not EOF(1)
    ctr = ctr + 1       'Adds one to the counter
    Input #1, CurrenciesNames(ctr), CurrenciesValue(ctr)        'Inputs two seperate arrays
Loop        'Loops back to the Do while not the end of file

For pass = 1 To ctr   'number of passes through the list
    For pos = 1 To ctr - 1  'Number of comparisons for each pass.
        If CurrenciesNames(pos) > CurrenciesNames(pos + 1) Then   'compare adjacent names
            tempName = CurrenciesNames(pos)  'swap if necessary
            CurrenciesNames(pos) = CurrenciesNames(pos + 1)
            CurrenciesNames(pos + 1) = tempName
            
            tempValue = CurrenciesValue(pos)        'Compare adjacent values
            CurrenciesValue(pos) = CurrenciesValue(pos + 1)     'Swap if necessary
            CurrenciesValue(pos + 1) = tempValue
        End If
    Next pos
Next pass
For X = 1 To ctr
    picResults.Print CurrenciesNames(X), , FormatCurrency(CurrenciesValue(X), 2)        'prints the newly sorted arrays in alphabetical order
Next X
Close #1        'Closes the file
End Sub

Private Sub cmdValue_Click()

Dim pass As Integer     'Sets variables
Dim pos As Integer
Dim tempName As String
Dim tempValue As Single
Dim X As Integer

pos = 0     'Initiates variables
pass = 0
tempName = 0
tempValue = 0
X = 0
ctr = 0
 
 picResults.Print "********************************World Currencies******************************************"
Open App.Path & "\Currencies.txt" For Input As #1       'Opens file as input

Do While Not EOF(1)
    ctr = ctr + 1       'Adds one to the counter
    Input #1, CurrenciesNames(ctr), CurrenciesValue(ctr)        'Inputs to seperate arrays
Loop

For pass = 1 To ctr   'number of passes through the list
    For pos = 1 To ctr - 1  'Number of comparisons for each pass.
        If CurrenciesValue(pos) > CurrenciesValue(pos + 1) Then   'compare adjacent values
            tempValue = CurrenciesValue(pos)        'Swap if necessary
            CurrenciesValue(pos) = CurrenciesValue(pos + 1)
            CurrenciesValue(pos + 1) = tempValue
            
            tempName = CurrenciesNames(pos)  'swap if necessary
            CurrenciesNames(pos) = CurrenciesNames(pos + 1)
            CurrenciesNames(pos + 1) = tempName
            
        End If
    Next pos
Next pass
For X = 1 To ctr
    picResults.Print CurrenciesNames(X), , FormatCurrency(CurrenciesValue(X), 2)        'Prints newly sorted array in order of monetary value
Next X
Close #1        'Closes the file
End Sub

Private Sub cmdClear_Click()

picResults.Cls      'Clears the picture box
End Sub

Private Sub cmdReturn_Click()

frmSorter.Hide      'Hides frmSorter
FrmMain.Show        'Shows frmMain
End Sub
