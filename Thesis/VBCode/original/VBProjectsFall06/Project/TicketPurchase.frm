VERSION 5.00
Begin VB.Form frmTicketPurchase 
   BackColor       =   &H00000080&
   Caption         =   "Ticket Purchase"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Card 
      Height          =   300
      ItemData        =   "TicketPurchase.frx":0000
      Left            =   7200
      List            =   "TicketPurchase.frx":0007
      MultiSelect     =   1  'Simple
      TabIndex        =   31
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Show view from seat"
      Height          =   735
      Left            =   2280
      TabIndex        =   30
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox txtNumber 
      Height          =   330
      Left            =   7200
      TabIndex        =   29
      Top             =   7440
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Index           =   0
      Left            =   7200
      TabIndex        =   28
      Top             =   6720
      Width           =   2895
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Purchase Tickets"
      Height          =   615
      Left            =   7800
      TabIndex        =   20
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Total"
      Height          =   615
      Left            =   7800
      TabIndex        =   19
      Top             =   3360
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   1335
      Left            =   7320
      ScaleHeight     =   1275
      ScaleWidth      =   2475
      TabIndex        =   13
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox txtSection 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox SeatingChart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   240
      Picture         =   "TicketPurchase.frx":001D
      ScaleHeight     =   4305
      ScaleWidth      =   3585
      TabIndex        =   6
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Enter Name:"
      Height          =   240
      Index           =   3
      Left            =   5880
      TabIndex        =   27
      Top             =   6720
      Width           =   1125
   End
   Begin VB.Label lblCredit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Credit Card:"
      Height          =   240
      Index           =   2
      Left            =   5880
      TabIndex        =   26
      Top             =   7080
      Width           =   1065
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Credit Card #:"
      Height          =   240
      Index           =   0
      Left            =   5880
      TabIndex        =   25
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Enter Billing Information:"
      Height          =   240
      Index           =   2
      Left            =   7800
      TabIndex        =   24
      Top             =   6240
      Width           =   2145
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Compute Total:"
      Height          =   240
      Index           =   1
      Left            =   7800
      TabIndex        =   23
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label lblStep4 
      BackColor       =   &H00000000&
      Caption         =   "Step 4:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   22
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblStep3 
      BackColor       =   &H00000000&
      Caption         =   "Step 3:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   21
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   "Pricing:"
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblPrice4 
      Caption         =   "Sections: 16-22    $30"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   17
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label lblPrice3 
      Caption         =   "Sections: 4-10    $30"
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   16
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblPrice2 
      Caption         =   "Sections: 11- 15 $25"
      Height          =   495
      Index           =   1
      Left            =   4080
      TabIndex        =   15
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblPrice1 
      Caption         =   "Sections 1-3, 23-24 $25"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   14
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblStep2 
      BackColor       =   &H00000000&
      Caption         =   "Step 2:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblStep1 
      BackColor       =   &H00000000&
      Caption         =   "Step 1:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblBuy 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Buy Gopher Hockey Tickets!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   6045
   End
   Begin VB.Label lblSec 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Section"
      Height          =   240
      Left            =   4440
      TabIndex        =   8
      Top             =   3000
      Width           =   825
   End
   Begin VB.Label lblSeats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Please select the best available seats by entering the section number below:"
      Height          =   240
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   6825
   End
   Begin VB.Label lblQt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Quantity"
      Height          =   240
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblPublic 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Public"
      Height          =   240
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Type of Ticket"
      Height          =   240
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Label lblQuantity 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Please enter the number of tickets you would like in the quantity box(es) below:"
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   6900
   End
End
Attribute VB_Name = "frmTicketPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmTicketPurchase
'Cole and John
'10/30/06
'Objective: The objective of this form is to allow the user to purcahse tickets
'for upcomming games.  The user is presented with a step by step process to do so.
'First, the user must input their desired quantity by inputting the quantity into
'the appropriate text box.  Secondly, the user inputs their desired section.  Third,
'the user computes their total by clicking on the compute total button. Finally, the
'user inputs their billing information by entering their name and credit card
'number into respective text boxes and selecting their credit card from a listbox.
'The user then enters that information by clicking on a command button entitled
'purchase tickets, and thus completed the process.

Option Explicit

Private Sub Card_DblClick()
    Card.Clear                          'this clears the list box
    Card.AddItem "Visa", 0                 'enters Visa as first selection
    Card.AddItem "Master Card", 1           'enters Master Card as second
    Card.AddItem "American Express", 2
    Card.AddItem "Discover", 3
End Sub

Private Sub cmdBack_Click()
    frmTicketPurchase.Visible = False
    frmTicketSales.Visible = True
End Sub


Private Sub cmdBuy_Click()
Dim Name As String
Dim Number As Long

    Name = txtName(0).Text
    Number = txtNumber.Text

MsgBox "Thank you " & Name & ". Your tickets have been purchased. CreditCard #" & Number & ""
End Sub

Private Sub cmdCompute_Click()
Dim Quantity, Section, Price As Integer
Dim Subtotal, Tax As Single, Total As Single

    Quantity = txtQuantity.Text             'input in Quantity textbox is Quantity variable
    Section = txtSection.Text
    
    If Quantity < 5 Then
        Quantity = Quantity
    Else
        MsgBox "Limit 4 tickets per purchase", , "Limit"    'sets a limit on how many tickets the user can buy
        Quantity = 0
    End If
    
    picResults.Cls
    
    Select Case Section             'if user selects a section, the appropriate price will be assesed
        Case 1 To 3
            Price = 25
        Case 4 To 10
            Price = 30
        Case 11 To 15
            Price = 25
        Case 16 To 22
            Price = 30
        Case 23 To 14
            Price = 25
        Case Else
            MsgBox "Section does not exist.", , "Error"
    End Select
    
    Subtotal = Quantity * Price         'computes subtotal
    Tax = Subtotal * 0.065                 'computes tax of 6.5%
    Total = Subtotal + Tax                  'computes total
    
    picResults.Print "Subtotal:", FormatCurrency(Subtotal)
    picResults.Print "Sales Tax:", FormatCurrency(Tax)
    picResults.Print "---------------------"
    picResults.Print "Total:", FormatCurrency(Total)
End Sub

Private Sub cmdView_Click()
    frmSectionView.Visible = True
    frmTicketPurchase.Visible = False
End Sub
