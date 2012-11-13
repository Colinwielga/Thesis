VERSION 5.00
Begin VB.Form frmStartup 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Wilderness Outfitter Inc."
   ClientHeight    =   8880
   ClientLeft      =   165
   ClientTop       =   945
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   Picture         =   "frmStartup.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   8640
   Begin VB.CommandButton cmdDays 
      Caption         =   "Enter"
      Height          =   495
      Left            =   6360
      TabIndex        =   30
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtDays 
      Height          =   285
      Left            =   4800
      TabIndex        =   27
      Text            =   "0"
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdInventory 
      Caption         =   "View Inventory"
      Height          =   615
      Left            =   5280
      TabIndex        =   26
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      Height          =   615
      Left            =   3600
      TabIndex        =   25
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdRent 
      Caption         =   "Rent"
      Height          =   615
      Left            =   1920
      TabIndex        =   24
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdGuide 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   4800
      Picture         =   "frmStartup.frx":128376
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdTowing 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   4800
      Picture         =   "frmStartup.frx":128783
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdTransportation 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   4800
      Picture         =   "frmStartup.frx":128D05
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowFile 
      Caption         =   "Show file"
      Height          =   615
      Left            =   240
      TabIndex        =   17
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6960
      TabIndex        =   16
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdMisc 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   360
      Picture         =   "frmStartup.frx":1292C3
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdEquip 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   360
      Picture         =   "frmStartup.frx":129888
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdBags 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   360
      Picture         =   "frmStartup.frx":129E26
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdTents 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   360
      Picture         =   "frmStartup.frx":12A4A0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdPacks 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   360
      Picture         =   "frmStartup.frx":12AB12
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCanoes 
      BackColor       =   &H00000080&
      Height          =   615
      Left            =   360
      Picture         =   "frmStartup.frx":12B0ED
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdBoats 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "frmStartup.frx":12B6DD
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   29
      Top             =   5160
      Width           =   615
   End
   Begin VB.Line Line8 
      X1              =   4680
      X2              =   7200
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line7 
      X1              =   7200
      X2              =   7200
      Y1              =   4440
      Y2              =   5640
   End
   Begin VB.Line Line6 
      X1              =   4680
      X2              =   7200
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line5 
      X1              =   4680
      X2              =   4680
      Y1              =   4440
      Y2              =   5640
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration of Rental:"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   28
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   8760
      Y2              =   7920
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8400
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line2 
      X1              =   8400
      X2              =   8400
      Y1              =   7920
      Y2              =   8760
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8400
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Guide Service"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   23
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Towing"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   20
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportation"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Camp Items Misc."
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rental Equipment"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Partial Outfitting"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sleeping Bags/Pads"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tents"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Canoes/Kayaks"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Packs"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Motor Boats"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Wilderness Outfitters"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wilderness Outfitters Partial Outfitting
'Justin Bork
'March, 2009
'
'frmStartup
'This form is the hub for all of the other forms.  It holds the buttons to
'access the rental forms and it also hold the exit button.  This form along with
'the many rental forms allow the user to browse rental options and design a
'trip.

 Private Sub cmdBags_Click()
    frmStartup.Hide
    frmSleep.Show
End Sub

Private Sub cmdBoats_Click()
    frmStartup.Hide
    frmBoats.Show
End Sub

Private Sub cmdCanoes_Click()
    frmStartup.Hide
    frmcanoes.Show
End Sub

Private Sub cmdDays_Click()
    'This button reads the number of days of the camping trip from a textbox
    'located in this form.  It then funtions the same as the other rental forms.
    'It too must reprint all items in order for the display to stay current.
    subtotal1 = 0
    subtotal2 = 0
    subtotal3 = 0
    grandTotal = 0
    
    frmDisplay.picResults.Cls
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
        grandTotal = grandTotal + (Subtotals(pos) * txtDays.Text)
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
    If txtDays.Text > 0 Then
        frmDisplay.picResults.Print "For a " & txtDays.Text & " day trip, Grand Total ="; Tab(40); FormatCurrency(grandTotal)
    End If

End Sub

Private Sub cmdEquip_Click()
    frmStartup.Hide
    frmEquip.Show
End Sub

Private Sub cmdGuide_Click()
    frmStartup.Hide
    frmGuide.Show
End Sub

Private Sub cmdInventory_Click()
    'Displays Inventory in frmInventory
    Dim InventoryItems(1 To 80) As String, Inventory(1 To 80) As Integer
    Dim InvCtr As Integer
    InvCtr = 0
    
    Open App.Path & "\inventory.txt" For Input As #1
    
    Do Until EOF(1)
        InvCtr = InvCtr + 1
        Input #1, InventoryItems(InvCtr), Inventory(InvCtr)
    Loop
    Close #1
    
    
    frmInventory.Show
    
    frmInventory.picResults.Print "Inventory Items"; Tab(35); "# Availible"
    frmInventory.picResults.Print
    For pos = 1 To 40
        frmInventory.picResults.Print InventoryItems(pos); Tab(40); Inventory(pos)
    Next pos
    
    frmInventory.picResults2.Print "Inventory Items"; Tab(35); "# Availible"
    frmInventory.picResults2.Print
    For pos = 41 To InvCtr
        frmInventory.picResults2.Print InventoryItems(pos); Tab(40); Inventory(pos)
    Next pos
End Sub

Private Sub cmdMisc_Click()
    frmStartup.Hide
    frmMisc.Show
End Sub

Private Sub cmdPacks_Click()
    frmStartup.Hide
    frmPacks.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRent_Click()
    'This button makes the rental selections official by removing them from
    'inventory.  It is also responsible for stoping the rental if an inventory
    'goes below zero.
    Dim InventoryItems(1 To 80) As String, Inventory(1 To 80) As Integer
    Dim InvCtr As Integer, pass As Integer
    Dim Found As Boolean, Found2 As Boolean
    InvCtr = 0
    Found = False
    Found2 = False
    pos = 0
    pass = 0
    
    Open App.Path & "\inventory.txt" For Input As #1
    
    Do Until EOF(1)
        InvCtr = InvCtr + 1
        Input #1, InventoryItems(InvCtr), Inventory(InvCtr)
    Loop
    Close #1
    
    Open App.Path & "\inventory.txt" For Output As #1
    
    For pos = 1 To InvCtr
        Inventory(pos) = Inventory(pos) - Requests(pos)
    Next pos
    
    For pos = 1 To InvCtr
        If Inventory(pos) < 0 Then
            frmOverdraw.Show
            frmOverdraw.picResults.Print InventoryItems(pos)
            Found = True
        End If
    Next pos
    
    If Found = True Then
        For pos = 1 To InvCtr
            Inventory(pos) = Inventory(pos) + Requests(pos)
            Write #1, InventoryItems(pos), Inventory(pos)
        Next pos
    ElseIf Found = False Then
        For pos = 1 To InvCtr
            Write #1, InventoryItems(pos), Inventory(pos)
        Next pos
        MsgBox "Have a great time! You are due back in " & txtDays.Text & " days!"
    End If
    
    Close #1
    
   
End Sub

Private Sub cmdReturn_Click()
    'This button returns all of the rented equipment to the inventories
    Dim InventoryItems(1 To 80) As String, Inventory(1 To 80) As Integer
    Dim InvCtr As Integer
    InvCtr = 0
    
    Open App.Path & "\inventory.txt" For Input As #1
    
    Do Until EOF(1)
        InvCtr = InvCtr + 1
        Input #1, InventoryItems(InvCtr), Inventory(InvCtr)
    Loop
    Close #1
    
    Open App.Path & "\inventory.txt" For Output As #1
    
    For pos = 1 To InvCtr
        Inventory(pos) = Inventory(pos) + Requests(pos)
        Write #1, InventoryItems(pos), Inventory(pos)
    Next pos
    
    Close #1
    
    MsgBox "Welcome back to civilization! We hope you had a great trip!"
End Sub

Private Sub cmdShowFile_Click()
    subtotal1 = 0
    subtotal2 = 0
    subtotal3 = 0
    grandTotal = 0
    
    'The following code prints all requested items under their respective
    'heading. Using this code in every rental form allows the display to always
    'be visible and up-to-date
    frmDisplay.picResults.Cls
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
        grandTotal = grandTotal + (Subtotals(pos) * txtDays.Text)
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
    If txtDays.Text > 0 Then
        frmDisplay.picResults.Print "For a " & txtDays.Text & " day trip, Grand Total ="; Tab(40); FormatCurrency(grandTotal)
    End If

End Sub

Private Sub cmdTents_Click()
    frmStartup.Hide
    frmTents.Show
End Sub

Private Sub cmdTowing_Click()
    frmStartup.Hide
    frmTowing.Show
End Sub

Private Sub cmdTransportation_Click()
    frmStartup.Hide
    frmTransport.Show
End Sub
