VERSION 5.00
Begin VB.Form frmGroceryStore 
   Caption         =   "Grocery Store"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   BeginProperty Font 
      Name            =   "NSimSun"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmGroceryStore.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmGroceryStore.frx":08CA
   ScaleHeight     =   8430
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here for Next"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Price"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort items by Price"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   2760
      Picture         =   "frmGroceryStore.frx":242D0C
      ScaleHeight     =   6555
      ScaleWidth      =   8115
      TabIndex        =   1
      Top             =   1560
      Width           =   8175
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Display your list of Groceries"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "$ Groceries Price List"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   7335
   End
End
Attribute VB_Name = "frmGroceryStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ingredients(1 To 30) As String, price(1 To 30) As Double, amount(1 To 30) As String
Dim I As Integer, CTR As Integer, CTR2 As Integer

Private Sub cmdDisplay_Click()

cmdTotal.Enabled = True
cmdPrice.Enabled = True

picResults.Cls

'Open the data file which select by the user
Open App.Path & groceryfile For Input As #1

'this loop reads data from a file into three arrays
CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, ingredients(CTR), price(CTR), amount(CTR)
Loop

picResults.Print "Ingredients"; Tab(30); "Price"; Tab(50); "Amount Per Serving"
picResults.Print "***************************************************************************************"
'this loop prints all runers and their names
For I = 1 To CTR
    picResults.Print ingredients(I); Tab(30); FormatCurrency(price(I)); Tab(50); amount(I)
Next I

Close #1

End Sub


Private Sub cmdNext_Click()
 
     frmGroceryStore.Hide
     frmLast.Show


End Sub

Private Sub cmdPrice_Click()

Dim pass As Integer, pos As Integer, Temp As Integer
Dim J As Integer, tempIngredients As String, tempPrice As Single, tempAmount


picResults.Cls

'use the Bubble sort to arrange the numbers in the desired order
'this code sorts the list of integer into ascending order
For pass = 1 To CTR - 1         'keep track of how many passes
    For pos = 1 To CTR - pass   'keep track of how many comparisons
        If price(pos) > price(pos + 1) Then
           tempIngredients = ingredients(pos)
           ingredients(pos) = ingredients(pos + 1)
           ingredients(pos + 1) = tempIngredients
           tempPrice = price(pos)
           price(pos) = price(pos + 1)
           price(pos + 1) = tempPrice
           tempAmount = amount(pos)
           amount(pos) = amount(pos + 1)
           amount(pos + 1) = tempAmount
        End If
    Next pos
Next pass

picResults.Print "Ingredients"; Tab(30); "Price"; Tab(50); "Amount Per Serving"
picResults.Print "***************************************************************************************"
'print the sorted list
    For J = 1 To CTR
        picResults.Print ingredients(J); Tab(30); FormatCurrency(price(J)); Tab(50); amount(J)
    Next J

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdTotal_Click()
Dim X As Double, Total As Double, CTR2 As Integer
picResults.Print
picResults.Print "Total Price(Round Off Number):"
picResults.Print "********************************"


X = 1
CTR2 = 0

For X = 1 To 30
    Total = Total + price(X)
    CTR2 = CTR2 + 1
    
Next
  picResults.Print FormatCurrency(Round(Total))
  
End Sub

