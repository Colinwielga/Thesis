VERSION 5.00
Begin VB.Form frmHandHeld 
   BackColor       =   &H00404080&
   Caption         =   "HandHeld Equipment"
   ClientHeight    =   9855
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDumbells 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Dumbells-$10.25"
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
      Left            =   3360
      Picture         =   "frmHandHeld.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   2655
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
      Height          =   975
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdHandball 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Hand Strengthing Balls-$2.99"
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
      Left            =   6360
      Picture         =   "frmHandHeld.frx":26AD
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton cmdAnkleWeights 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ankle Weights-$9.95"
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
      Left            =   360
      Picture         =   "frmHandHeld.frx":3146
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdMat 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Floor Mat-     $14.95"
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
      Left            =   360
      Picture         =   "frmHandHeld.frx":4A60
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   2535
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
      Height          =   975
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdDone 
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
      Height          =   975
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Height          =   975
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      Height          =   4575
      Left            =   3600
      ScaleHeight     =   4515
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Your Equipment"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11415
   End
End
Attribute VB_Name = "frmHandHeld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Build You Own Home Gym
'Form Name: frmHandHeld
'Author: Michelle Pickle
'Date Written: March 12th 2009
'The objective of this form is to allow the user to purchase various handheld workout equipment to begin to build their home gym.
    'This can be achieved by clicking on the picture of the object.  This allows the user to visualize what they are building
    'The user is asked the quantity he/she wishes to purchase
Option Explicit
'forces the user to declare all of his/her variables

Private Sub cmdAnkleWeights_Click()
'the variables are declared
Dim Pair As Integer
Dim Subtotal As Double
Dim pairtotal As Double
'input box asks the user to enter the quanitiy of weights they would like and is then set equal to "Pair"
    Pair = InputBox("Please Enter the Number of Ankle Weights Desiered", "Ankle Weights")
'calculates the price by mulitply the number of pairs requested(got from user) by the cost of the ankle weights.
    Subtotal = Pair * 9.95
    handheldtotal = handheldtotal + Subtotal
'add the cost of the ankle weights to the running total.  Running total is declared in the modules so it is avaiable to all forms
    runningtotal = runningtotal + handheldtotal
'displays the item name and the total price in money format
    picResults.Print "Ankle Weights"; Tab(20); Pair; Tab(30); FormatCurrency(Subtotal)
    
End Sub

Private Sub cmdDone_Click()
'this button allows the user to advance to the next form
    frmHandHeld.Hide
    frmReceipt.Show
End Sub

Private Sub cmdDumbells_Click()
'variables declared
Dim Dumbell As Integer
Dim Subtotal As Double
Dim Dumbelltotal As Double
'asks the user how many sets of dumbells he/she wants and then assigns that number to "dumbell"
    Dumbell = InputBox("How many Dumbells would you like?", "Dumbell")
'calculates the price of the desired dumbells
    Subtotal = Dumbell * 10.25
    handheldtotal = handheldtotal + Subtotal
'add the total of the dumbells to the runningtotal (aka total of everything the user selects)
    runningtotal = runningtotal + handheldtotal
'prints the item name and the total price with dollar sign
    picResults.Print "Dumbells"; Tab(20); Dumbell; Tab(30); FormatCurrency(Subtotal)
    
    
End Sub

Private Sub cmdHandball_Click()
'variables declared
Dim Hand As Integer
Dim Subtotal As Double
Dim handtotal As Double
'asks the user how many hand strenghtening balls he/she wants and then assigns the entered number to "hand"
    Hand = InputBox("Please Enter the Number Hand Strengthing Balls Desiered", "Strength Balls")
'multiplies number entered by user by the prices of handballs
    Subtotal = Hand * 2.99
    handheldtotal = handheldtotal + Subtotal
'total of all the purchases
    runningtotal = runningtotal + handheldtotal
'prints the item name and total price of that item
    picResults.Print "Stregth Balls"; Tab(20); Hand; Tab(30); FormatCurrency(Subtotal)
    
End Sub

Private Sub cmdMat_Click()
'variables declared
Dim Mat As Integer
Dim Subtotal As Double
Dim mattotal As Double
'gets number of mats the user wants from the user and assigns the entered number to "Mat"
    Mat = InputBox("Please Enter the Number of Mats Desired", "Mats")
'calates the total cost of the mat(s) requested
    Subtotal = Mat * 14.95
    handheldtotal = handheldtotal + mattotal
'add the total of the mat to the overall total
    runningtotal = runningtotal + Subtotal
'prints the item name and the cost, with dollar sign
    picResults.Print "Mat"; Tab(20); Mat; Tab(30); FormatCurrency(Subtotal)
End Sub

Private Sub cmdNext_Click()
'changes forms
    frmHandHeld.Hide
    frmMachines.Show
End Sub

Private Sub cmdQuit_Click()
'ends the program
    End
End Sub

Private Sub cmdTotal_Click()
'This button lets the user know his/her total so far (every item purchased, on this screen, is taken into account)
    picResults.Print "******************************************************"
    picResults.Print "Your total for handheld equipment is "; FormatCurrency(handheldtotal, 2)
End Sub


Private Sub Form_Load()
'This code centers the form on computer screen upon loading.
'this code discovered from Cassie Scherer and Jordan Schmaltz project of developing a vacation

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
