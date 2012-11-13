VERSION 5.00
Begin VB.Form frmGiftShop 
   BackColor       =   &H00FF0000&
   Caption         =   "Gift Shop"
   ClientHeight    =   8505
   ClientLeft      =   2310
   ClientTop       =   1500
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10785
   Begin VB.PictureBox pic1 
      BackColor       =   &H008080FF&
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7875
      ScaleWidth      =   2715
      TabIndex        =   14
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton cmdTickets 
         BackColor       =   &H0000C000&
         Caption         =   "Buy Your Tickets Now"
         Height          =   975
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmdTop 
         BackColor       =   &H000080FF&
         Caption         =   "Top 10 Disney Animated Movies Of All Time"
         Height          =   855
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton cmdTrivia 
         BackColor       =   &H0000FFFF&
         Caption         =   "Trivia Game "
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdIntro 
         BackColor       =   &H000000FF&
         Caption         =   "Main Page "
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00800080&
         Caption         =   "Quit "
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6840
         Width           =   1815
      End
   End
   Begin VB.PictureBox picMugs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8640
      Picture         =   "frmGiftShop.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picStuffedAnimals 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5880
      Picture         =   "frmGiftShop.frx":0D93
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox picMickeyMouse 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4920
      Picture         =   "frmGiftShop.frx":1B8E
      ScaleHeight     =   1755
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox picActionFigures 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9360
      Picture         =   "frmGiftShop.frx":2804
      ScaleHeight     =   1755
      ScaleWidth      =   1155
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picTShirts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3120
      Picture         =   "frmGiftShop.frx":32B0
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdMugs 
      BackColor       =   &H0080FF80&
      Caption         =   "Disney Mugs $6"
      Height          =   855
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdStuffedAnimals 
      BackColor       =   &H0080FF80&
      Caption         =   "Stuffed Animals $25"
      Height          =   975
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdMickeyMouse 
      BackColor       =   &H0080FF80&
      Caption         =   "Mickey Mouse Ears  $5"
      Height          =   1215
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdActionFigures 
      BackColor       =   &H0080FF80&
      Caption         =   "Action Figures $8"
      Height          =   975
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdTshirts 
      BackColor       =   &H0080FF80&
      Caption         =   "T-Shirts $15"
      Height          =   975
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear"
      Height          =   300
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H0080FF80&
      Caption         =   "Total"
      Height          =   300
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FF00&
      Height          =   3255
      Left            =   3000
      ScaleHeight     =   3195
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   4680
      Width           =   7695
   End
   Begin VB.Label lblGiftShop 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Your Shopping Bag"
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   4200
      Width           =   2655
   End
End
Attribute VB_Name = "frmGiftShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RunningTotal As Single
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form to allow users to purchase merchandise from the Disney Gift Shop and display the total (including tax).

Private Sub cmdActionFigures_Click()
Dim ActionFigures As Single
ActionFigures = 8       'amount per item
picResults.Print " Figures "; , FormatCurrency(ActionFigures) 'format currency displays the numbers in monetary form.
End Sub

Private Sub cmdClear_Click()
picResults.Cls          'Allows user to clear
RunningTotal = RunningTotal * 0 'resets running total
End Sub

Private Sub cmdKeyChains_Click()
Dim KeyChains As Single
KeyChains = 7       'amount per item
RunningTotal = RunningTotal + KeyChains     'Keeps a running total of all the different items bought at the gift shop.
picResults.Print " Key Chains "; , FormatCurrency(KeyChains)
End Sub




Private Sub cmdGiftShop_Click()
frmGiftShop.Show        'Allows user to go to the Gift Shop form
frmTrivia.Hide
frmIntro.Hide
frmTop.Hide
frmTickets.Hide

End Sub

Option Explicit

Private Sub cmdIntro_Click()
frmGiftShop.Hide        'Allows user to go to the Intro form
frmTrivia.Hide
frmIntro.Show
frmTop.Hide
frmTickets.Hide

End Sub

Private Sub cmdMickeyMouse_Click()
Dim MickeyMouse As Single
MickeyMouse = 5
RunningTotal = RunningTotal + MickeyMouse   'Keeps total running
picResults.Print " Mickey Ears "; , FormatCurrency(MickeyMouse)
End Sub

Private Sub cmdMugs_Click()
Dim Mugs As Single
Mugs = 6
RunningTotal = RunningTotal + Mugs  'Keeps total running
picResults.Print " Mugs "; , FormatCurrency(Mugs)
End Sub


Private Sub cmdQuit_Click() 'Allows user to quit the program
End
End Sub

Private Sub cmdStuffedAnimals_Click()
Dim StuffedAnimals As Single
StuffedAnimals = 25     'Amount per item
RunningTotal = RunningTotal + StuffedAnimals        'Keeps total running
picResults.Print " Animals "; , FormatCurrency(StuffedAnimals)
End Sub

Private Sub cmdTickets_Click()
frmGiftShop.Hide        'Allows user to go to the Ticket form
frmIntro.Hide
frmTrivia.Hide
frmTop.Hide
frmTickets.Show

End Sub

Private Sub cmdTop_Click()
frmGiftShop.Hide            'Allows user to go to the Top ten movies form
frmIntro.Hide
frmTrivia.Hide
frmTop.Show
frmTickets.Hide

End Sub

Private Sub cmdTotal_Click()
Dim Total As Single
Dim Tax As Single
Tax = RunningTotal * 0.05   'Multiplies the total times the tax
picResults.Print , "--------------------------"
picResults.Print "subtotal:", FormatCurrency(RunningTotal) 'runningtotal before tax
picResults.Print "Tax:", FormatCurrency(Tax) 'just .05 * runningtotal
picResults.Print "Total:", FormatCurrency(RunningTotal + Tax) 'adds runningtotal and tax together and then prints them out.
End Sub
Private Sub cmdTrivia_Click()
frmGiftShop.Hide        'Allows user to go to the Trivia form
frmIntro.Hide
frmTrivia.Show
frmTop.Hide
frmTickets.Hide

End Sub

Private Sub cmdTshirts_Click()
Dim Tshirts As Single
Tshirts = 15        'amount per item
RunningTotal = RunningTotal + Tshirts   'keeps total running
picResults.Print " T-Shirts "; , FormatCurrency(Tshirts)
End Sub




