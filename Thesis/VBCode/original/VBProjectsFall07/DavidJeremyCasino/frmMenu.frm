VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Menu"
   ClientHeight    =   7695
   ClientLeft      =   2925
   ClientTop       =   1935
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   8520
   Begin VB.CommandButton cmdLeave 
      BackColor       =   &H000040C0&
      Caption         =   "Back To Lobby"
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C000C0&
      Caption         =   "Clear Purchases"
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000C0C0&
      Height          =   7695
      Left            =   4320
      ScaleHeight     =   7635
      ScaleWidth      =   4155
      TabIndex        =   6
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdTotal 
         BackColor       =   &H0000C000&
         Caption         =   "Total Purchase"
         Height          =   855
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5760
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdSalad 
      Height          =   1935
      Left            =   0
      Picture         =   "frmMenu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdHotDog 
      Height          =   1935
      Left            =   2280
      Picture         =   "frmMenu.frx":930A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdHamburger 
      Height          =   1935
      Left            =   0
      Picture         =   "frmMenu.frx":1712C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdSundae 
      Height          =   1935
      Left            =   2280
      Picture         =   "frmMenu.frx":1F378
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdGrilledCheese 
      Height          =   1935
      Left            =   0
      Picture         =   "frmMenu.frx":27BF4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSteak 
      Height          =   1935
      Left            =   2280
      Picture         =   "frmMenu.frx":35F7A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   7725
      Left            =   0
      Picture         =   "frmMenu.frx":4367F
      Top             =   0
      Width           =   4320
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This form provides with the option to buy food
Dim total As Single
Dim clicker As Boolean

Private Sub cmdClear_Click()
    'This button clears total
    clicker = False
    total = 0
    picResults.Cls
    cmdTotal.Enabled = True
    cmdSundae.Enabled = True
    cmdHamburger.Enabled = True
    cmdHotDog.Enabled = True
    cmdSteak.Enabled = True
    cmdSalad.Enabled = True
    cmdGrilledCheese.Enabled = True
End Sub

Private Sub cmdGrilledCheese_Click()
    'When clicked 1 Grilled Cheese is added to running total
    clicker = True
    picResults.Print "Grilled Cheese"; Tab(20); FormatCurrency(2)
    total = total + 2
End Sub

Private Sub cmdHamburger_Click()
    'When clicked 1 Hamburger is added to running total
    clicker = True
    picResults.Print "Hamburger"; Tab(20); FormatCurrency(3.5)
    total = total + 3.5
End Sub

Private Sub cmdHotDog_Click()
    'When clicked 1 Hot Dog is added to running total
    clicker = True
    picResults.Print "Hot Dog"; Tab(20); FormatCurrency(1.5)
    total = total + 1.5
End Sub

Private Sub cmdLeave_Click()
    'Go back to Lobby
    frmMenu.Hide
    frmLobby.Show
End Sub

Private Sub cmdSalad_Click()
    'When clicked 1 Salad is added to running total
    clicker = True
    picResults.Print "Salad"; Tab(20); FormatCurrency(4)
    total = total + 4
End Sub

Private Sub cmdSteak_Click()
    'When clicked 1 Steak is added to running total
    clicker = True
    picResults.Print "Steak"; Tab(20); FormatCurrency(20)
    total = total + 20
End Sub

Private Sub cmdSundae_Click()
    'When clicked 1 Sundae is added to running total
    clicker = True
    picResults.Print "Sundae"; Tab(20); FormatCurrency(3)
    total = total + 3
End Sub

Private Sub cmdTotal_Click()
    'Adds the food clicked with tax and displays it in a picturebox
    Dim totaltotal As Single
    Dim tax As Single
    tax = total * 0.07
    totaltotal = total + tax
    If totaltotal < balanceglobal Then
        balanceglobal = balanceglobal - totaltotal
        If clicker = True Then
            picResults.Print "******************************************"
            picResults.Print "Subtotal"; Tab(20); FormatCurrency(total)
            picResults.Print "Tax"; Tab(20); FormatCurrency((total * 0.07))
            picResults.Print "******************************************"
            picResults.Print "Total"; Tab(20); FormatCurrency(totaltotal)
        End If
    Else
        MsgBox "Get more money fool!", , "Show me the money"
    End If
    cmdTotal.Enabled = False
    cmdSundae.Enabled = False
    cmdHamburger.Enabled = False
    cmdHotDog.Enabled = False
    cmdSteak.Enabled = False
    cmdSalad.Enabled = False
    cmdGrilledCheese.Enabled = False

End Sub

