VERSION 5.00
Begin VB.Form frmNikeBasketball 
   BackColor       =   &H80000007&
   Caption         =   "NikeBasketball"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   Picture         =   "frmNikeBasketball.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBacktoHome 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Store Home"
      Height          =   1455
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   1455
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back To Nike Store"
      Height          =   1455
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortNu 
      BackColor       =   &H000000FF&
      Caption         =   "Sort by Cost"
      Height          =   1455
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortbyName 
      BackColor       =   &H000000FF&
      Caption         =   "Sort Alphabetically"
      Height          =   1455
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdReadData 
      BackColor       =   &H000000FF&
      Caption         =   "Read Data"
      Height          =   1455
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   5295
      Left            =   7680
      ScaleHeight     =   5235
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmNikeBasketball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
' NikeBasketball
' Andrew Eisinger
' 3/20/09
'This program reads data set
'This program prints data
'This program sorts alphabetically
'This program sorts by price

Private Sub cmdBack_Click()
frmNike1.Show
frmNikeBasketball.Hide
End Sub

Private Sub cmdBacktoHome_Click()
frmStoreHome.Show
frmNikeBasketball.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReadData_Click()
picResults.Cls
Open App.Path & "\NikeBasketball.txt" For Input As #1
CTR = 0
picResults.Print "Basketball Item"; Tab(40); "Basketball Costs"
picResults.Print "**************************************************************************"
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, BasketballItems(CTR), BasketballCosts(CTR)
    picResults.Print BasketballItems(CTR); Tab(40); FormatCurrency(BasketballCosts(CTR))
Loop
Close #1
End Sub

Private Sub cmdSortbyName_Click()
'This button sorts the arrays by the first names of the products
Dim J As Single, TempBasketballItems As String, TempBasketballCosts As Single
    'Clears the previous results
    picResults.Cls
    
    'Prints headings so it is easier for the user to understand
    'the results.
    picResults.Print "Basketball Item"; Tab(40); "Basketball Cost"
    picResults.Print "**************************************************************************"
 
    
    'Code to sort the two parralel arrays by the first name of the product
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If BasketballItems(Pos) > BasketballItems(Pos + 1) Then
                TempBasketballItems = BasketballItems(Pos)
                BasketballItems(Pos) = BasketballItems(Pos + 1)
                BasketballItems(Pos + 1) = TempBasketballItems
                TempBasketballCosts = BasketballCosts(Pos)
                BasketballCosts(Pos) = BasketballCosts(Pos + 1)
                BasketballCosts(Pos + 1) = TempBasketballCosts
            End If
        Next Pos
    Next Pass
    For J = 1 To CTR
        
        'Prints the results
        picResults.Print BasketballItems(J); Tab(40); FormatCurrency(BasketballCosts(J))
    Next
End Sub

Private Sub cmdSortNu_Click()
Dim J As Single, TempBasketballItems As String, TempBasketballCosts As Single
    'Clears the previous results
    picResults.Cls
    
    'Prints headings so it is easier for the user to understand
    'the results.
    picResults.Print "Basketball Item"; Tab(40); "Basketball Cost"
    picResults.Print "**************************************************************************"
    
    
    'Code to sort the two parralel arrays by the price of the product
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If BasketballCosts(Pos) > BasketballCosts(Pos + 1) Then
                TempBasketballCosts = BasketballCosts(Pos)
                BasketballCosts(Pos) = BasketballCosts(Pos + 1)
                BasketballCosts(Pos + 1) = TempBasketballCosts
                TempBasketballItems = BasketballItems(Pos)
                BasketballItems(Pos) = BasketballItems(Pos + 1)
                BasketballItems(Pos + 1) = TempBasketballItems
            End If
        Next Pos
    Next Pass
    For J = 1 To CTR
        
        'Prints the results
        picResults.Print BasketballItems(J); Tab(40); FormatCurrency(BasketballCosts(J))
    Next
End Sub
