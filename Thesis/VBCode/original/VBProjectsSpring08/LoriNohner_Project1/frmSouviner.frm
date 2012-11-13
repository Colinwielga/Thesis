VERSION 5.00
Begin VB.Form frmSouviner 
   BackColor       =   &H0000C000&
   Caption         =   "Buy a Souviner"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00FFFF00&
      Caption         =   "Total Cost"
      Height          =   855
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00FFFF00&
      Caption         =   "Reset List "
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdDust 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Magic Pixie Dust:   $24.99"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdGlobe 
      BackColor       =   &H00FF80FF&
      Caption         =   "Disney Castle Snow Globe:                      $9.99"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdSword 
      BackColor       =   &H008080FF&
      Caption         =   "Hero's Sword:        $14.99"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdCrown 
      BackColor       =   &H0080FFFF&
      Caption         =   "Princess Crown:     $14.99"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdHat 
      BackColor       =   &H00FF8080&
      Caption         =   "Mickey Mouse Hat: $19.99"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   4080
      ScaleHeight     =   4035
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit"
      Height          =   255
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFF00&
      Caption         =   "Return to Disney Castle"
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "frmSouviner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Project
'Souviner
'Lori Nohner
'Written- March 17, 2008
'Objective- Asks the user how much they want to spend on souviners.
    'The user can click on the souviners they want to buy and the price is displayed in the picture box
    'the total is added up and tax is added.  If the user spent to much a message box is displayed.
Option Explicit
Dim Total As Single

Private Sub cmdCrown_Click()
    Dim Crown As Single 'declares variable as single
    Crown = 14.99 'sets price of item
    Total = Total + Crown 'adds item to total price
    picResults.Print "Princess Crown"; Tab(35); FormatCurrency(Crown) 'prints out name and cost of item
End Sub

Private Sub cmdDust_Click()
    Dim Dust As Single 'declares variable as single
    Dust = 24.99 'sets item cost
    Total = Total + Dust 'adds item to total cost
    picResults.Print "Magic Pixie Dust"; Tab(35); FormatCurrency(Dust) 'prints out name and cost of item
End Sub

Private Sub cmdExit_Click()
    End 'quits program
End Sub

Private Sub cmdGlobe_Click()
    Dim Globe As Single 'declares variable as single
    Globe = 9.99 'sets item cost
    Total = Total + Globe 'adds item to total cost
    picResults.Print "Disney Castle Snow Globe"; Tab(35); FormatCurrency(Globe) 'prints out name and cost of item
End Sub

Private Sub cmdHat_Click()
    Dim Hat As Single 'declares variable as single
    Hat = 19.99 'sets item cost
    Total = Total + Hat 'adds item to total cost
    picResults.Print "Mickey Mouse Hat"; Tab(35); FormatCurrency(Hat) 'prints out name and cost of item
    
End Sub

Private Sub cmdReset_Click()
 'resets total and clears picture box
 picResults.Cls
 
 Total = 0
End Sub

Private Sub cmdReturn_Click()
    frmSouviner.Hide 'hides souvenir page
    frmDisneyCastle.Show 'returns to Disney Home Page
    
End Sub

Private Sub cmdSword_Click()
    Dim Sword As Single 'declares variable as single
    Sword = 14.99 'sets item cost
    Total = Total + Sword 'adds item to total cost
    picResults.Print "Hero's Sword"; Tab(35); FormatCurrency(Sword) 'prints out name and cost of item
End Sub

Private Sub cmdTotal_Click()
    Dim Tax As Single 'declares variable as single
    Dim TotalCost As Single 'declares variable as single
    Tax = 0.07 * Total ' formula for tax
    TotalCost = Tax + Total 'formula for total cost
    
    'if the amount that the user wanted to spend is greater then the amount spent a message box appears and the total soct is reset to zero
    If TotalCost > Money Then
        MsgBox "Oops. You bought too much.", , "Error"
        picResults.Cls
        Total = 0
    
    Else
        picResults.Print ' prints balnk line
        picResults.Print ' prints balnk line
        picResults.Print "Item Total"; Tab(35); FormatCurrency(Total) 'prints total item cost
        picResults.Print "Tax"; Tab(35); FormatCurrency(Tax) 'prints amount of tax
        picResults.Print ' prints blank line
        picResults.Print "Total Cost"; Tab(35); FormatCurrency(TotalCost) 'prints total item cost with tax
    End If
    
End Sub

