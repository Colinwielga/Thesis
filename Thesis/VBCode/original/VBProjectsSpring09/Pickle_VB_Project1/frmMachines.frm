VERSION 5.00
Begin VB.Form frmMachines 
   BackColor       =   &H000000C0&
   Caption         =   "Machines"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Previous Screen"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton cmdLastScreen 
      BackColor       =   &H00FFC0FF&
      Caption         =   "I'm Done Buying"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdBike 
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   7920
      Picture         =   "frmMachines.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton cmdWeightMachine 
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4320
      Picture         =   "frmMachines.frx":191B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   2775
   End
   Begin VB.CommandButton cmdElliptical 
      Caption         =   "Elliptical-     $350.50"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   720
      Picture         =   "frmMachines.frx":3974
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton cmdTreadmill 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Treadmill-$295.99"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   720
      Picture         =   "frmMachines.frx":58B2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   5415
      Left            =   4320
      ScaleHeight     =   5355
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label lblBike 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Stationary Bike-$195.95"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   10200
      Width           =   2655
   End
   Begin VB.Label lblWeightMachine 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Weight Machine-  $220.95"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   10200
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Machines Available For Purchase"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   12855
   End
End
Attribute VB_Name = "frmMachines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Build Your Own Home Gym
'Form Name: frmMachines
'Author: Michelle Pickle
'Date Written: March 12th 2009
'The purpose of this form is allow the user to purchase bigger machines in order to fulfill cardio, or get their heart moving.
    'Once again, the picture help the user to visualize what they are "buying"
Option Explicit
'forces the user to declare all of his/her variables in order for the program to run
'when a button is pressed, the user is essentailly purchasing the product

Private Sub cmdBike_Click()
Dim Subtotal As Double
Dim biketotal As Double
'calculates the cost of the bike and adds the cost to the running total
    biketotal = 195.95
    equipmenttotal = equipmenttotal + biketotal
    runningtotal = runningtotal + equipmenttotal
'prints the item name and the item price
    picResults.Print "Stationary Bike"; Tab(20); FormatCurrency(biketotal)
End Sub

Private Sub cmdElliptical_Click()
Dim Subtotal As Double
Dim Ellipticaltotal As Double
'caculates the cost of an elliptical including tax
'adds the cost of the elliptical to the overall total
    Ellipticaltotal = 350.5
    equipmenttotal = equipmenttotal + Ellipticaltotal
    runningtotal = runningtotal + equipmenttotal
'prints the item name and the item price
    picResults.Print "Elliptical"; Tab(20); FormatCurrency(Ellipticaltotal)
End Sub

Private Sub cmdLastScreen_Click()
'this changes forms, going to final screen
    frmMachines.Hide
    frmReceipt.Show
End Sub

Private Sub cmdNext_Click()
'this button changes forms, allowing the user to advance to the next form
    frmMachines.Hide
    frmGyms.Show
End Sub

Private Sub cmdPrevious_Click()
'this button allows the user to return to the previous form
    frmHandHeld.Visible = True
    frmMachines.Visible = False
End Sub

Private Sub cmdQuit_Click()
'ends the program
    End
End Sub

Private Sub cmdTotal_Click()
'prints the current overall total for all items purchased in correct money format
    picResults.Print "************************************"
    picResults.Print "Your total for equipment is "; FormatCurrency(equipmenttotal)
End Sub

Private Sub cmdTreadmill_Click()
    Dim Subtotal As Double
    Dim Treadmilltotal As Double
'calculates the cost of the treadmill including tax
'adds the cost of the treadmill to the overall cost of the items
    Treadmilltotal = 295.99
    equipmenttotal = equipmenttotal + Treadmilltotal
    runningtotal = runningtotal + equipmenttotal
'prints the item name and the item price
    picResults.Print "Treadmill"; Tab(20); FormatCurrency(Treadmilltotal)

End Sub


Private Sub cmdWeightMachine_Click()
Dim Subtotal As Double
Dim Weighttotal As Double
'calculates the cost of the weight machine
'adds the price of the item to the overal price for the individuals purchases
    Weighttotal = 220.95
    equipmenttotal = equipmenttotal + Weighttotal
    runningtotal = runningtotal + Weighttotal
    
'prints the item name and the price of the item
    picResults.Print "Weight Machine"; Tab(20); FormatCurrency(Weighttotal)
End Sub

Private Sub Form_Load()
'This code centers the form on computer screen upon loading.
'this code discovered from Cassie Scherer and Jordan Schmaltz project of developing a vacation

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
