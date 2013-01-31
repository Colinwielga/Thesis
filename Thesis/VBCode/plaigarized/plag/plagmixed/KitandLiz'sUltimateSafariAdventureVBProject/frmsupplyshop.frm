VERSION 5.00
Begin VB.Form frmsupplyshop 
   BackColor       =   &H0000C0C0&
   Caption         =   "Safari Supplies Shop"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      FillColor       =   &H00C0C0FF&
      Height          =   5895
      Left            =   2280
      ScaleHeight     =   5835
      ScaleWidth      =   4515
      TabIndex        =   10
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton cmdreturntohomepage 
      Caption         =   "Go to Safari HQ"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturntoJungle 
      Caption         =   "Return to Jungle"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Cmdquit 
      Caption         =   " I Quit! Get Me Home!"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdTourGuide 
      Caption         =   "Add a Safari Tour Guide!"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdTransportation 
      Caption         =   "Transportation"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSafariGear 
      Caption         =   "Safari Gear"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdWaterandfood 
      Caption         =   "Water and Food"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdCamera 
      Caption         =   "Camera"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   495
      Left            =   7200
      OleObjectBlob   =   "frmsupplyshop.frx":0000
      SourceDoc       =   "M:\CS130\Kit and Liz's Ultimate Safari Adventure VB Project\Unknown Artist\Unknown Album (2-24-2010 6-02-05 PM)\01 Track 1 (2).wma"
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblsong 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Click here to listen to the song Kit would play if she lived on the Safari... or anywhere for that matter!"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1335
      Left            =   7080
      TabIndex        =   12
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblSafariShop 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   " Welcome to the Safari shop! We hope you find everything you need for your ultimate Safari Adventure!"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   975
      Left            =   2280
      TabIndex        =   11
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmsupplyshop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
         
'Declares variables
Dim runningtotal As Double

Private Sub cmdCamera_Click()
'Declares more variables
Dim Camera As Double
 'Says the price of the camera and adds it to the total
Camera = 45#
runningtotal = runningtotal + Camera
'Prints the price of th Camera in the correct form, dollars and cents
picResults.Print "Camera"; Tab(20); FormatCurrency(Camera)
End Sub

Private Sub cmdClear_Click()
'Clears the picture box
picResults.Cls

End Sub

Private Sub cmdQuit_Click()
'Quit button
End

End Sub

Private Sub cmdreturntohomepage_Click()
'Brings you back to the home page
frmsupplyshop.Hide
 'The Great Safari Adventure
  'Frm The Supply Shope
  'Kit and Liz Chambers
  'February 21st 2010
  'Objective: The purpose of this form is to
         'Print the Price of objects needed for a safari adventure
         'Calculate the total
         'Add the tax
FrmWelcome.Show

End Sub

Private Sub cmdReturntoJungle_Click()
'Brings you to the jungle page
frmTheJungle.Show
frmsupplyshop.Hide

End Sub

Private Sub cmdSafariGear_Click()
'Adds the price of safari gear to the running total and calcuates it in dollars and cents
Dim SafariGear As Double
SafariGear = 100#
runningtotal = runningtotal + SafariGear + 1 - 1
'Prints the Safari gear
picResults.Print "Safari Gear"; Tab(20); FormatCurrency(SafariGear)
End Sub

Private Sub cmdTotal_Click()
 'Declares Variables
Dim Total As Double
Dim Subtotal As Double
Dim Tax As Double
'Prints a row of stars
picResults.Print "***********************************"

'Calculates the total
Subtotal = runningtotal
picResults.Print "SubTotal"; Tab(20); FormatCurrency(Subtotal)
'Calcualtes and adds the tax
Tax =  runningtotal * 0.07
'Prints the Tax
picResults.Print "Tax"; Tab(20); FormatCurrency(Tax)
Total = Tax + Subtotal
'Prints the Total
picResults.Print "Total"; Tab(20); FormatCurrency(Total)

End Sub

Private Sub cmdTourGuide_Click()
'Declares Variables
Dim Tourguide As Double
 'Displays and adds the price of the tour guide to the running total
Tourguide = 95#
runningtotal = runningtotal + Tourguide
picResults.Print "Tour Guide"; Tab(20); FormatCurrency(Tourguide)
End Sub

Private Sub cmdTransportation_Click()
 'Declares Variables
Dim Transportation As Double
 'Displays the price of transportation
Transportation = 30#
 'Adds it to the runningTotal
runningtotal = Transportation + runningtotal
picResults.Print "Transportation"; Tab(20); FormatCurrency(Transportation)
End Sub

Private Sub cmdWaterandfood_Click()
 'Declares Variables
Dim WaterandFood As Double
 'Adds the price of the food and water to the running total
WaterandFood = 15#
runningtotal = WaterandFood + runningtotal
picResults.Print "Water and Food"; Tab(20); FormatCurrency(WaterandFood)
End Sub
