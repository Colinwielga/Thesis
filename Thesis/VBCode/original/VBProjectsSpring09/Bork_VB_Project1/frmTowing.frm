VERSION 5.00
Begin VB.Form frmTowing 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Wilderness Outfitters"
   ClientHeight    =   6225
   ClientLeft      =   1455
   ClientTop       =   945
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   Picture         =   "frmTowing.frx":0000
   ScaleHeight     =   6225
   ScaleWidth      =   7320
   Begin VB.CommandButton cmdPreviousTowing 
      Caption         =   "Previous Page"
      Height          =   495
      Left            =   2280
      TabIndex        =   24
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubmitTowing 
      Caption         =   "Submit"
      Height          =   495
      Left            =   480
      TabIndex        =   23
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtTowing 
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   5
      Text            =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtTowing 
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   4
      Text            =   "0"
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtTowing 
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   3
      Text            =   "0"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtTowing 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   2
      Text            =   "0"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtTowing 
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   1
      Text            =   "0"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtTowing 
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Text            =   "0"
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$67.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   22
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$62.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label15 
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
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$24.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$32.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   5640
      TabIndex        =   16
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "To Prairie Portage to Back Bay Portage"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   4080
      Width           =   3795
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "To Prairie Portage to  Washington Island"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   3600
      Width           =   3795
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "To Birch Lake or Prairie Portage"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   3120
      Width           =   3795
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Moose Lake to Ensign Portage"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   2640
      Width           =   3795
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Vermilion to Trout Lake Portge"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2160
      Width           =   3795
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Burntside to Crab Lake Portage"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   1680
      Width           =   3795
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "(of people)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Towing"
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
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmTowing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wilderness Outfitters Partial Outfitting
'Justin Bork
'March, 2009
'
'frmTowing
'The purpose of this form is to allow the user to get a tow (boat ride) to
'various available destinations.
Private Sub cmdPreviousTowing_Click()
    frmTowing.Hide
    frmStartup.Show
End Sub

Private Sub cmdSubmitTowing_Click()
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
    
    For pos = 87 To 92
        i = i + 1
        Requests(pos) = txtTowing(i).Text
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
