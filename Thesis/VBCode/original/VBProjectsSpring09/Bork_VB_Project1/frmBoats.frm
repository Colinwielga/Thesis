VERSION 5.00
Begin VB.Form frmBoats 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Wilderness Outfitters"
   ClientHeight    =   6675
   ClientLeft      =   2865
   ClientTop       =   945
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   Picture         =   "frmBoats.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   5775
   Begin VB.TextBox txtBoats 
      Height          =   285
      Index           =   9
      Left            =   360
      TabIndex        =   32
      Text            =   "0"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtBoats 
      Height          =   285
      Index           =   8
      Left            =   360
      TabIndex        =   31
      Text            =   "0"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtBoats 
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   30
      Text            =   "0"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtBoats 
      Height          =   285
      Index           =   6
      Left            =   360
      TabIndex        =   29
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtBoats 
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   28
      Text            =   "0"
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtBoats 
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   27
      Text            =   "0"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtBoats 
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   26
      Text            =   "0"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtBoats 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   25
      Text            =   "0"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtBoats 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   24
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmdReturnBoats 
      Caption         =   "Previous Page"
      Height          =   615
      Left            =   2640
      TabIndex        =   23
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdSubmitBoats 
      Caption         =   "Submit"
      Height          =   615
      Left            =   360
      TabIndex        =   22
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "No."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$3.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$8.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$10.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$40.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$30.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$25.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$28.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$28.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$26.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Unit Rate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rental Items"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Boat Seat"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Portage Wheels"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Electric Trolling Motor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Outboard Motor 25 hp"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Outboard Motor 15 hp"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Outboard Motor 8 hp"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "19' Sqare Stern Canoe"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "16' Fishing Boat"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "14' Fishing Boat"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Boats and Motors / Related Equipment"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmBoats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wilderness Outfitters Partial Outfitting
'Justin Bork
'March, 2009
'
'frmBoats
'The purpose of this form is to allow the user to choose from the boating items
'available for rental.

Private Sub cmdReturnBoats_Click()
   frmBoats.Hide
   frmStartup.Show
End Sub
Private Sub cmdSubmitBoats_Click()
    'This button first takes note of what items are requested and writes that
    'into the user's text file.
    'It then prints the user's selections.
    
    Dim i As Integer
    subtotal1 = 0
    subtotal2 = 0
    subtotal3 = 0
    grandTotal = 0
    i = 0
    
    frmDisplay.picResults.Cls
    
    For pos = 1 To 9
        i = i + 1
        Requests(pos) = txtBoats(i).Text
        Subtotals(pos) = Prices(pos) * Requests(pos)
    Next pos
    
    Open App.Path & "\Customers\" & user & "\" & user & year & ".txt" For Output As #1
    
    For pos = 1 To counter
        Write #1, Items(pos), Prices(pos), Requests(pos), Subtotals(pos)
    Next pos
    
    Close #1
    'The following code prints all requested items under their respective
    'heading. Using this code in every rental form allows the display to always
    'be visible and up-to-date
    
    frmDisplay.picResults.Print Tab(40); "Prices"; Tab(50); "Number"; Tab(60); "Subtotals"
    frmDisplay.picResults.Print "--Motor Boats--"
    
    For pos = 1 To 9
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Canoes/Kayaks--"
   
    For pos = 10 To 27
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Rental Equipment--"
 
    For pos = 28 To 37
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Packs--"

    For pos = 38 To 42
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Tents--"
   
    For pos = 43 To 46
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Sleeping Bags/Pads--"

    For pos = 47 To 51
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Camp Items Misc.--"
 
    For pos = 52 To 64
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Transportation--"

    For pos = 65 To 86
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Towing--"

    For pos = 87 To 92
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Guide Service--"

    If Requests(93) > 0 Then
        frmDisplay.picResults.Print Items(93); Tab(40); FormatCurrency(Prices(93)); Tab(50); Requests(93); Tab(60); FormatCurrency(Subtotals(93))
    End If
    
    'The following code first adds up the seperate subtotals of the rental itmes
    'and saves the values in certain groups so that they can be used for
    'calculating taxes. The subtotals and taxes are then added up along with a
    'certain group of items that are multiplies by the number of days in the
    'trip. The user ends up with a grand total for the trip.
    For pos = 1 To 64
        subtotal1 = subtotal1 + Subtotals(pos)
    Next pos
    
    For pos = 65 To counter
        subtotal2 = subtotal2 + Subtotals(pos)
    Next pos
    
    For pos = 87 To counter
        subtotal3 = subtotal3 + Subtotals(pos)
    Next pos
    
    For pos = 1 To 63
        grandTotal = grandTotal + (Subtotals(pos) * frmStartup.txtDays.Text)
    Next pos
    
    salesTax = subtotal1 * 0.065
    lodgingTax = Subtotals(64) * 0.03
    fsTax = subtotal3 * 0.03
    total = subtotal1 + subtotal2 + salesTax + lodgingTax + fsTax
    grandTotal = grandTotal + Subtotals(64) + subtotal2 + salesTax + lodgingTax + fsTax
    
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "Subtotal:"; Tab(25); FormatCurrency(subtotal1 + subtotal2)
    frmDisplay.picResults.Print "Sales Tax (6.5%):"; Tab(25); FormatCurrency(salesTax)
    frmDisplay.picResults.Print "Lodging Tax (3%):"; Tab(25); FormatCurrency(lodgingTax)
    frmDisplay.picResults.Print "USFS Tax (3%):"; Tab(25); FormatCurrency(fsTax)
    frmDisplay.picResults.Print "------------------------------------------------------"
    frmDisplay.picResults.Print "Total:"; Tab(25); FormatCurrency(total)
    If frmStartup.txtDays.Text > 0 Then
        frmDisplay.picResults.Print "For a " & frmStartup.txtDays.Text & " day trip, Grand Total ="; Tab(40); FormatCurrency(grandTotal)
    End If
End Sub


