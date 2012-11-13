VERSION 5.00
Begin VB.Form frmBakery 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10800
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10800
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Click Here to Sort Items by Price (lowest to highest)"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   4215
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FF8080&
      Caption         =   "Calculate Bakery Subtotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   3855
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Calculate Cost Of This Item and Add to Cart"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   3855
   End
   Begin VB.TextBox txtNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Click Here to See Available Items and Prices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   11280
      ScaleHeight     =   4995
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00FF00FF&
      Caption         =   "Continue Shopping or Check Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label lblNumber 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter How Many You Would Like"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   6
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label lblItem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Item Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   5025
      Left            =   360
      Picture         =   "frmBakery.frx":0000
      Top             =   120
      Width           =   5985
   End
   Begin VB.Label lblWelcomeBakery 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the Bakery Section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmBakery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR As Integer
Dim BakedGoods(1 To 50) As String
Dim Prices(1 To 50) As Single
Dim CostOfItem As Single

Private Sub cmdAdd_Click()

picResults.Print "**********************************************"
picResults.Print "Bakery Subtotal: "; FormatCurrency(BakeryRunningTotal)
picResults.Print "**********************************************"

End Sub

Private Sub cmdCalculate_Click()
Dim Item As Integer, Number As Integer
'this multiplies the item by the cost of the item as derived from the array
Number = txtNumber.Text
Item = txtItem.Text

CostOfItem = Number * Prices(Item)
'this displays the results
picResults.Print "*********************************"
picResults.Print "Item Cost: "; FormatCurrency(CostOfItem)
picResults.Print "*********************************"

'this adds the cost of the item to the running total of all bakery item
BakeryRunningTotal = BakeryRunningTotal + CostOfItem

End Sub

Private Sub cmdContinue_Click()
'this adds the bakery subtotal to the runningtotal for all sections

RunningTotal = BakeryRunningTotal + ProduceRunningTotal + FrozenRunningTotal
'this displays the amount to the user
MsgBox "Total spent so far is: " & FormatCurrency(RunningTotal)
'this navigates the user back to the enter form
frmBakery.Hide
frmProduce.Hide
frmFrozen.Hide
frmCheckOut.Hide
frmEnter.Show

End Sub

Private Sub cmdLoad_Click()
Dim Pos As Integer

CTR = 0

picResults.Cls

'this opens a data file with bakery items and prices

Open App.Path & "\BakeryPrices.txt" For Input As #1

Do Until EOF(1)
   CTR = CTR + 1
   Input #1, BakedGoods(CTR), Prices(CTR)
Loop
Close #1


picResults.Print "Item #"; Tab(10); "Item Name"; Tab(35); "Price"
picResults.Print "*************************************************************"

'this displays the data in a picture box

For Pos = 1 To CTR
    picResults.Print Pos; " : "; Tab(10); BakedGoods(Pos); Tab(35); FormatCurrency(Prices(Pos))
Next Pos

cmdLoad.Enabled = False
cmdSort.Enabled = True
  
End Sub


Private Sub cmdSort_Click()
Dim Pass As Integer
Dim Temp As Single
Dim Temp2 As String
Dim Pos As Integer

picResults.Cls

picResults.Print "Item #"; Tab(10); "Item Name"; Tab(35); "Price"
picResults.Print "*************************************************************"
'this sorts the bakery items by price from lowest to highest

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Prices(Pos) > Prices(Pos + 1) Then
            Temp = Prices(Pos)
            Prices(Pos) = Prices(Pos + 1)
            Prices(Pos + 1) = Temp
            Temp2 = BakedGoods(Pos)
            BakedGoods(Pos) = BakedGoods(Pos + 1)
            BakedGoods(Pos + 1) = Temp2
        End If
    Next Pos
Next Pass
'this displays the results
For Pos = 1 To CTR
    picResults.Print Pos; " : "; Tab(10); BakedGoods(Pos); Tab(35); FormatCurrency(Prices(Pos))
Next Pos

End Sub
