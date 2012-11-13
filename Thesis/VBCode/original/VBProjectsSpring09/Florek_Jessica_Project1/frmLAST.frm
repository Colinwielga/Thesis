VERSION 5.00
Begin VB.Form frmLAST 
   BackColor       =   &H00FF8080&
   Caption         =   "Form2"
   ClientHeight    =   10260
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9225
   LinkTopic       =   "Form2"
   Picture         =   "frmLAST.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBudget 
      Caption         =   "Show Budget"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdFinished 
      Caption         =   "I am Finished!"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0FFC0&
      FillColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   3720
      ScaleHeight     =   9915
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
   End
End
Attribute VB_Name = "frmLAST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmLAST
'Jessica Florek
'Written: 3/21/09
'Objective: This form shows the users overall trip expenses and budget information
'in addition to calculating their travel expense.

Option Explicit
Dim j As Integer, travelexpense(1 To 10) As Single


Private Sub cmdBack_Click()
picResults.Cls
'so the next time this is accessed the form will print from the top of the picture box
frmLAST.Hide
frmMapCities.Show

End Sub

Private Sub cmdBudget_Click()
'summary of overall budget information
picResults.Print "Amount Budgeted "; Tab(35); FormatCurrency(budget2)
picResults.Print "*********************************"
picResults.Print "Remaining Budget "; Tab(35); FormatCurrency(budget)
picResults.Print "Approximate Cost of Food"; Tab(35); FormatCurrency(foodcost)
picResults.Print "Flight Cost:"; Tab(35); FormatCurrency(flightcost)
picResults.Print

j = 0

'opens approximate travel costs, in an array so that they can be changed as eurorail, bus, or flight prices change
Open App.Path & "\TravelCosts.txt" For Input As #1
Do While Not EOF(1)
    j = j + 1
    Input #1, travelexpense(j)
Loop
Close #1

picResults.Print "Travel Expenses"
picResults.Print "*********************************"

'this is the estimated costs of travel. case statement because the message(type of travel) varies based on number of cities visited
Select Case citycounter
Case Is >= 1
    picResults.Print "Cost of Bus/Taxi/Subway Travel"; Tab(35); FormatCurrency(travelexpense(citycounter))
Case Is = 2
    picResults.Print "Cost of Traveling between Two Cities"; Tab(35); FormatCurrency(travelexpense(citycounter))
Case Is = 3
    picResults.Print "Cost of Traveling between Three Cities"; Tab(35); FormatCurrency(travelexpense(citycounter))
Case Is = 4
    picResults.Print "Cost of Traveling between Four Cities"; Tab(35); FormatCurrency(travelexpense(citycounter))
Case Is > 4
    picResults.Print "An error has occured."
End Select

picResults.Print

'each of the following will only be displayed IF the city has been visited
If london = True Then
    picResults.Print "London Costs"
    picResults.Print "-------------------------------------------"
    picResults.Print "Hotel Expense"; Tab(35); FormatCurrency(Londonhotelcost)
    picResults.Print "Entertainment Expense"; Tab(35); FormatCurrency(londonattractioncost)
    picResults.Print
End If

If paris = True Then
    picResults.Print "Paris Costs"
    picResults.Print "-------------------------------------------"
    picResults.Print "Hotel Expense"; Tab(35); FormatCurrency(Parishotelcost)
    picResults.Print "Entertainment Expense"; Tab(35); FormatCurrency(parisattractioncost)
    picResults.Print
End If

If venice = True Then
    picResults.Print "Venice Costs"
    picResults.Print "-------------------------------------------"
    picResults.Print "Hotel Expense"; Tab(35); FormatCurrency(Venicehotelcost)
    picResults.Print "Entertainment Expense"; Tab(35); FormatCurrency(veniceattractioncost)
    picResults.Print
End If

If madrid = True Then
    picResults.Print "Madrid Costs"
    picResults.Print "-------------------------------------------"
    picResults.Print "Hotel Expense"; Tab(35); FormatCurrency(Madridhotelcost)
    picResults.Print "Entertainment Expense"; Tab(35); FormatCurrency(madridattractioncost)
    picResults.Print
End If

picResults.Print "Total Expenses"
picResults.Print "*********************************"
picResults.Print "Total Hotel Expense"; Tab(35); FormatCurrency(Madridhotelcost + Londonhotelcost + Venicehotelcost + Parishotelcost)
picResults.Print "Total Entertainment Expense"; Tab(35); FormatCurrency(madridattractioncost + londonattractioncost + veniceattractioncost + parisattractioncost)
picResults.Print

picResults.Print "Total Cost of Trip"; Tab(35); FormatCurrency(budget2 + Abs(budget))



End Sub

Private Sub cmdFinished_Click()
MsgBox ("Thank You!")
End
End Sub
