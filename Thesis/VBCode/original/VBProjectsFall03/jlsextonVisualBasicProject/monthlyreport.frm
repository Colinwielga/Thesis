VERSION 5.00
Begin VB.Form MonthlyReport 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Monthly Report"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture6 
      Height          =   2895
      Left            =   120
      Picture         =   "monthlyreport.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   915
      TabIndex        =   12
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      Height          =   2655
      Left            =   1560
      Picture         =   "monthlyreport.frx":0CC0
      ScaleHeight     =   2595
      ScaleWidth      =   915
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      Height          =   2775
      Left            =   3000
      Picture         =   "monthlyreport.frx":16FA
      ScaleHeight     =   2715
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   0
      Picture         =   "monthlyreport.frx":28AB
      ScaleHeight     =   3555
      ScaleWidth      =   4155
      TabIndex        =   9
      Top             =   480
      Width           =   4215
   End
   Begin VB.PictureBox Picture3 
      Height          =   2895
      Left            =   8880
      Picture         =   "monthlyreport.frx":54E3
      ScaleHeight     =   2835
      ScaleWidth      =   915
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   7800
      Picture         =   "monthlyreport.frx":67DD
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdprofitability 
      Caption         =   "Most Profitable Product"
      Height          =   855
      Left            =   7800
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdinventory 
      Caption         =   "Inventory Remaining"
      Height          =   855
      Left            =   6720
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox results 
      Height          =   2055
      Left            =   4320
      ScaleHeight     =   1995
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   2040
      Width           =   4455
   End
   Begin VB.CommandButton cmdprofit 
      Caption         =   "Profit"
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   8880
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Load 
      Caption         =   "May Monthly Report"
      Height          =   855
      Left            =   4320
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   "   May Finacial Report"
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "MonthlyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'project name : WakeboardingShop (Monthlyreport.vbp)
'Form name : Monthly Report (MonthlyReport.frm)
'Author : Jim Sexton
'Date Written: October 21, 2003
'Purpose of the Form: To allow the user to obtain certain finacial
                    'information about a company using some key figures
                    'like the product name, price, cost, and quantity sold,
                    'the program will be able to calculate revenue, expenses
                    'profit, and the inventory remaining in
                    'stock of a certain product.  Allowing the user
                    'to input a product name and print out the quantity. The program
                    'also is able to sort the most profitable product.

Option Explicit
Dim strpath As String       'Global variaables that will be used throughout the program are dimed here
Dim strname(1 To 4) As String
Dim price(1 To 4) As Single
Dim cost(1 To 4) As Single
Dim quantitysold(1 To 4) As Integer
Dim revenue(1 To 4) As Single
Dim profit(1 To 4) As Single
Dim expenses(1 To 4) As Single
Dim totalrevenue As Single
Dim totalexpenses As Single
Dim totalprofit As Single

Private Sub cmdinventory_Click()
Dim found As Boolean              'Variables used just for this button were dimed here
Dim i As Integer
Dim n As String
Dim inventory(1 To 4) As Integer
found = False
results.Cls                        'Closes the array if one was previously opened
n = InputBox("Please Enter Product", "Product") 'User is asked for what type of product is left in stock
i = 0
Do Until i = 4 Or found = True     'The program searches each product to find if there is a match
        i = i + 1
    If n = strname(i) Then
        found = True
    End If
Loop
If found = True Then                            'If the product is found then program prints out what product and the quantity remaining in stock
        inventory(i) = 150 - quantitysold(i)
        results.Print n; " has"; inventory(i); "remaining in stock"
    Else
        results.Print n; " is not in inventory " 'If the product is not found then message is printing stating the product is not in stock
End If
End Sub

Private Sub cmdprofitability_Click()
Dim strname2(1 To 4) As String
Dim profit2(1 To 4) As Single
Dim i As Integer                  'Variables are dimed for this command button only
Dim n As Integer
Dim temp As String
Dim tempp As Single
Dim pass As Integer
n = 4
results.Cls
For i = 1 To 4                  'this makes a copy of the profit and name so it doesnt change the array
    profit2(i) = profit(i)
    strname2(i) = strname(i)
Next i
results.Print "Product", "Profit"
results.Print "--------------------------------------------"
For pass = 1 To n - 1                'The copied array is sorted by the most profitable product
    For i = 1 To n - pass
        If profit2(i) < profit2(i + 1) Then
                tempp = profit2(i)
                profit2(i) = profit2(i + 1)
                profit2(i + 1) = tempp
                temp = strname2(i)
                strname2(i) = strname2(i + 1)
                strname2(i + 1) = temp
        End If
    Next i
Next pass
For i = 1 To 4
    results.Print strname2(i), FormatCurrency(profit2(i))   'Prints the results starting with the most profitable product to the least profitable product, both product name and its profit
Next i
    results.Print "--------------------------------------------"
    results.Print "Total", FormatCurrency(totalprofit)      'The total profit is printed
End Sub

Private Sub cmdquit_Click()
End                                 'The end botton command will end the program
End Sub

Private Sub cmdprofit_Click()
Dim i As Integer
results.Cls
    results.Print "Product", "Revenue", "Expenses", "Profit"   'Pinting headings for the form
    results.Print "--------------------------------------------------------------------------------------------"
For i = 1 To 4
    results.Print strname(i), FormatCurrency(revenue(i)), FormatCurrency(expenses(i)), FormatCurrency(profit(i)) 'Prints the revenues expenses and profits the program prints each products revenue, expense, and profit
Next i
    results.Print "--------------------------------------------------------------------------------------------"
    results.Print "Total", FormatCurrency(totalrevenue), FormatCurrency(totalexpenses), FormatCurrency(totalprofit) 'Prints the total revene, total expenses, and total profit
End Sub

Private Sub Form_Load()
Dim i As Integer
strpath = "N:\CS130\handin\jlsexton Visual Basic Project\data.txt"    'File is opened and information from file is loaded into an array
Open strpath For Input As #1
    For i = 1 To 4
         Input #1, strname(i), price(i), cost(i), quantitysold(i) 'Inputs all the information from file
    Next i
For i = 1 To 4
    revenue(i) = price(i) * quantitysold(i)           'Using the information from the file the program calculates revenue, expenses, and profit
    expenses(i) = cost(i) * quantitysold(i)
    profit(i) = revenue(i) - expenses(i)
    totalrevenue = totalrevenue + revenue(i)
    totalexpenses = totalexpenses + expenses(i)
    totalprofit = totalprofit + profit(i)
Next i
Close #1
End Sub

Private Sub Load_Click()
Dim i As Integer        'Variable is dimed for this command button only
results.Cls
Dim month As String
    results.Print , "May Report"  'Prints out headings for the form
    results.Print , "------------------------"
    results.Print "Product", "Price", "Product Cost", "Quantity Sold"
    results.Print "--------------------------------------------------------------------------------------------"
For i = 1 To 4
    results.Print strname(i), FormatCurrency(price(i)), FormatCurrency(cost(i)), quantitysold(i)  'Prints out the information
Next i
End Sub

