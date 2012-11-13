VERSION 5.00
Begin VB.Form frmSalad 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   8280
      TabIndex        =   13
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   8280
      TabIndex        =   12
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   5880
      TabIndex        =   11
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   735
      Left            =   5880
      TabIndex        =   10
      Top             =   5520
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   8280
      ScaleHeight     =   4035
      ScaleWidth      =   1995
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalad8 
      Caption         =   "Fish-$8"
      Height          =   1815
      Left            =   720
      Picture         =   "frmSalad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalad7 
      Caption         =   "Peppers-$4"
      Height          =   1815
      Left            =   3240
      Picture         =   "frmSalad.frx":0EF5
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalad6 
      Caption         =   "Eggplant-$3"
      Height          =   1815
      Left            =   5760
      Picture         =   "frmSalad.frx":23F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalad5 
      Caption         =   "Green Beans-$2"
      Height          =   1935
      Left            =   5760
      Picture         =   "frmSalad.frx":357B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalad4 
      Caption         =   "Cabages-$2"
      Height          =   1815
      Left            =   3240
      Picture         =   "frmSalad.frx":467C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalad3 
      Caption         =   "Olives-$5"
      Height          =   1935
      Left            =   3240
      Picture         =   "frmSalad.frx":5645
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalad2 
      Caption         =   "Cucumbers-$4"
      Height          =   1815
      Left            =   720
      Picture         =   "frmSalad.frx":6655
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalad1 
      Caption         =   "Lotus-$3"
      Height          =   1935
      Left            =   720
      Picture         =   "frmSalad.frx":75E0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      Caption         =   "Choose the salad by clikcing the buttons"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmSalad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Digital Menu
'Form Name: frmSalad
'Authors: Gaole Chen
'Date Written: 3/7/09
'Objective: The user can order the salad by clicking the vivid pictures.
'The form also calculate the total price automatically for the user.

Option Explicit
Dim runningTotal As Single

Private Sub cmdBack_Click()
frmMenu.Show
frmSalad.Hide
End Sub

Private Sub cmdClear_Click()
picResults.Cls
End Sub

Private Sub cmdNext_Click()
frmMain.Show
frmSalad.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSalad1_Click()
'declare the variables
Dim Numberone As Integer, Totalone As Integer
'calculate how many Lotus the user wants
Numberone = InputBox("How many Lotus would you like?")
Totalone = 3 * Numberone
runningTotal = runningTotal + Totalone
'show the user the amount and price
picResults.Print Numberone; " Lotus:", FormatCurrency(Totalone)

End Sub

Private Sub cmdSalad2_Click()
'declare the variables
Dim Numberfive As Integer, Totalfive As Integer
'calculate how many cucumbers the user wants
Numberfive = InputBox("How many cucumbers would you like?")
Totalfive = 4 * Numberfive
runningTotal = runningTotal + Totalfive
'show the user the amount and price
picResults.Print Numberfive; " Cucumbers:"; FormatCurrency(Totalfive)
End Sub

Private Sub cmdSalad3_Click()
'declare the variables
Dim Numbertwo As Integer, Totaltwo As Integer
'calculate how many Olives the user wants
Numbertwo = InputBox("How many Olives would you like?")
Totaltwo = 5 * Numbertwo
runningTotal = runningTotal + Totaltwo
'show the user the amount and price
picResults.Print Numbertwo; " Olives:", FormatCurrency(Totaltwo)
End Sub

Private Sub cmdSalad4_Click()
'declare the variables
Dim Numbersix As Integer, Totalsix As Integer
'calculate how many Cabages the user wants
Numbersix = InputBox("How many Cabages would you like?")
Totalsix = 2 * Numbersix
runningTotal = runningTotal + Totalsix
'show the user the amount and price
picResults.Print Numbersix; " Cabages:", FormatCurrency(Totalsix)
End Sub

Private Sub cmdSalad5_Click()
'declare the variables
Dim Numberthree As Integer, Totalthree As Integer
'calculate how many Green Beans the user wants
Numberthree = InputBox("How many Green Beans would you like?")
Totalthree = 2 * Numberthree
runningTotal = runningTotal + Totalthree
'show the user the amount and price
picResults.Print Numberthree; " Green Beans:"; FormatCurrency(Totalthree)
End Sub

Private Sub cmdSalad6_Click()
'declare the variables
Dim Numberseven As Integer, Totalseven As Integer
'calculate how many Eggplants the user wants
Numberseven = InputBox("How many Eggplants would you like?")
Totalseven = 3 * Numberseven
runningTotal = runningTotal + Totalseven
'show the user the amount and price
picResults.Print Numberseven; " Eggplant:", FormatCurrency(Totalseven)
End Sub

Private Sub cmdSalad7_Click()
'declare the variables
Dim Numberfour As Integer, Totalfour As Integer
'calculate how many Peppers the user wants
Numberfour = InputBox("How many Peppers would you like?")
Totalfour = 4 * Numberfour
runningTotal = runningTotal + Totalfour
'show the user the amount and price
picResults.Print Numberfour; " Peppers:", FormatCurrency(Totalfour)
End Sub

Private Sub cmdSalad8_Click()
'declare the variables
Dim Numbereight As Integer, Totaleight As Integer
'calculate how many fish the user wants
Numbereight = InputBox("How many Fish would you like?")
Totaleight = 8 * Numbereight
runningTotal = runningTotal + Totaleight
'show the user the amount and price
picResults.Print Numbereight; " Fish:", FormatCurrency(Totaleight)
End Sub

Private Sub cmdTotal_Click()
'declare the variables
Dim Taxrate As Single, Tax As Single
picResults.Print "-------------------------------------------------"
Taxrate = 0.08
'calculate the total price of the salad
Tax = runningTotal * Taxrate
runningTotal = runningTotal + Tax
picResults.Print "Taxes:", FormatCurrency(Tax)
picResults.Print "The salad you ordered"
picResults.Print "in total costs: "; FormatCurrency(runningTotal); "."
Totalsaladcost = Totalsaladcost + runningTotal
Totalcost = Totalcost + Totalsaladcost
End Sub

