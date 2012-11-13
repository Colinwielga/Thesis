VERSION 5.00
Begin VB.Form frmgear 
   BackColor       =   &H00400040&
   Caption         =   "Form1"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17520
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   17520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsort 
      Caption         =   "Sort By Cost"
      Height          =   855
      Left            =   13440
      TabIndex        =   21
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   15240
      TabIndex        =   20
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton cmdbuycoat 
      Caption         =   "Buy"
      Height          =   735
      Left            =   8160
      TabIndex        =   19
      Top             =   7440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox piccoat 
      Height          =   2655
      Left            =   5160
      Picture         =   "frmgear.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   18
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdbuyhoodiebear 
      Caption         =   "Buy"
      Height          =   735
      Left            =   8160
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pichoodiebear 
      Height          =   3135
      Left            =   4920
      Picture         =   "frmgear.frx":13F8
      ScaleHeight     =   3075
      ScaleWidth      =   2475
      TabIndex        =   16
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdbuymonopoly 
      Caption         =   "Buy"
      Height          =   615
      Left            =   2760
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picmonopoly 
      Height          =   2295
      Left            =   360
      Picture         =   "frmgear.frx":2E44
      ScaleHeight     =   2235
      ScaleWidth      =   2115
      TabIndex        =   14
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdbuypurse 
      Caption         =   "Buy"
      Height          =   615
      Left            =   2760
      TabIndex        =   13
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPurse 
      Height          =   1335
      Left            =   720
      Picture         =   "frmgear.frx":6790
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdbuyfootball 
      Caption         =   "Buy"
      Height          =   735
      Left            =   8160
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next"
      Height          =   855
      Left            =   13560
      TabIndex        =   10
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "Total"
      Height          =   855
      Left            =   11760
      TabIndex        =   9
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton cmdclearpic 
      Caption         =   "Clear"
      Height          =   855
      Left            =   15240
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdloadprices 
      Caption         =   "Load Prices"
      Height          =   855
      Left            =   11760
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox picfootball 
      Height          =   2295
      Left            =   4560
      Picture         =   "frmgear.frx":6F19
      ScaleHeight     =   2235
      ScaleWidth      =   3195
      TabIndex        =   6
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdbuypatatoe 
      Caption         =   "Buy"
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picmrpatatoe 
      Height          =   2055
      Left            =   360
      Picture         =   "frmgear.frx":9515
      ScaleHeight     =   1995
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdbuyjersey 
      Caption         =   "Buy"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picjersey 
      Height          =   2175
      Left            =   360
      Picture         =   "frmgear.frx":A929
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   7815
      Left            =   11640
      ScaleHeight     =   7755
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Label lblgear 
      BackColor       =   &H00400040&
      Caption         =   "Every Brett Favre Super Fan Needs Gear. Click the buttons to add the items to your cart:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11175
   End
End
Attribute VB_Name = "frmgear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdbuycoat_Click()
'the following are same for each item button. buyclicking the button the user is purchasing the item which is being kept track of in the picresults box and in the running total
Dim coat As Single
coat = 98
picResults.Print "Vikings Coat"; Tab(40); FormatCurrency(coat)
runningtotal = runningtotal + coat
End Sub

Private Sub cmdbuyfootball_Click()
Dim football As Single
football = 26
picResults.Print "Vikings Football"; Tab(40); FormatCurrency(football)
runningtotal = runningtotal + football
End Sub

Private Sub cmdbuyhoodiebear_Click()
Dim bear As Single
bear = 18
picResults.Print "Vikings Hoodie Bear"; Tab(40); FormatCurrency(bear)
runningtotal = runningtotal + bear
End Sub

Private Sub cmdbuyjersey_Click()
Dim jersey As Single
jersey = 150
picResults.Print "Jersey"; Tab(40); FormatCurrency(jersey)
runningtotal = runningtotal + jersey
End Sub

Private Sub cmdbuymonopoly_Click()
Dim game As Single
game = 20
picResults.Print "Vikings Monopoly Game"; Tab(40); FormatCurrency(game)
runningtotal = runningtotal + game
End Sub

Private Sub cmdbuypatatoe_Click()
Dim mr As Single
mr = 39
picResults.Print "Vikings Mr. Potatoe Head"; Tab(40); FormatCurrency(mr)
runningtotal = runningtotal + mr
End Sub

Private Sub cmdbuypurse_Click()
Dim purse As Single
purse = 35
picResults.Print "Vikings Purse"; Tab(40); FormatCurrency(purse)
runningtotal = runningtotal + purse
End Sub

Private Sub cmdclearpic_Click()
picResults.Cls
runningtotal = 0
End Sub
'Project Name: Brett Favre Fan Club
'Form Name: frmgear
'Author: Kory LaCroix
'Date Written: 10/19/08
'Objective: To buy fan gear
Private Sub cmdloadprices_Click()
CTR = 0
'This will load the following file
Open App.Path & "\gear.txt" For Input As #1
'This will place the information in the file into two parrallel arrays
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, item(CTR), itemcost(CTR)
Loop

picResults.Print "Item"; Tab(40); "Cost"
picResults.Print "******************************************************************"
'This will display the information from the file
For j = 1 To CTR
    picResults.Print item(j); Tab(40); FormatCurrency(itemcost(j))
Next j
'This will hide the load prices button. This prevents users from clicking on the button twice.
cmdloadprices.Visible = False

End Sub

Private Sub cmdnext_Click()
'This will go the next form
frmgear.Hide
frmschedule.Show
End Sub

Private Sub cmdQuit_Click()
'this will end the program
End
End Sub

Private Sub cmdsort_Click()
'These are the names for the bubble sort
Dim Tempitem As String
Dim Tempitemcost As Double
Dim Pass As Integer

picResults.Cls

picResults.Print " "
picResults.Print "Item"; Tab(40); "Item Cost"
picResults.Print "************************************************"
'This will sort the gear from highest price to lowest price
    For Pass = 1 To CTR - 1
    For j = 1 To CTR
        If itemcost(j) < itemcost(j + 1) Then
            Tempitem = item(j)
            item(j) = item(j + 1)
            item(j + 1) = Tempitem
            Tempitemcost = itemcost(j)
            itemcost(j) = itemcost(j + 1)
            itemcost(j + 1) = Tempitemcost
        End If
    Next j
    Next Pass

    For j = 1 To CTR
        'this is what prints the sorted items
        picResults.Print item(j); Tab(40); FormatCurrency(itemcost(j))
    Next j
    
picResults.Print " "
picResults.Print "Item"; Tab(40); "Item Cost"
picResults.Print "************************************************"

'This will enable the user to now click on the buttons to purchase the items.
cmdbuyjersey.Visible = True
cmdbuypatatoe.Visible = True
cmdbuymonopoly.Visible = True
cmdbuypurse.Visible = True
cmdbuyfootball.Visible = True
cmdbuyhoodiebear.Visible = True
cmdbuycoat.Visible = True
End Sub

Private Sub cmdtotal_Click()
Dim tax As Double
Dim total As Double
Dim taxtotal As Double
Dim subtotal As Double

picResults.Print " "
picResults.Print "The total of your Breatt Favre Fan Gear is:"
picResults.Print "*********************************************************"
'this will create a total and then calculate the tax and then give a final total
picResults.Print "Subtotal = "; Tab(40); FormatCurrency(runningtotal)

tax = 0.065
taxtotal = tax * runningtotal
total = taxtotal + runningtotal
picResults.Print "Tax= "; Tab(40); FormatCurrency(taxtotal)
picResults.Print "Total= "; Tab(40); FormatCurrency(total)

End Sub
