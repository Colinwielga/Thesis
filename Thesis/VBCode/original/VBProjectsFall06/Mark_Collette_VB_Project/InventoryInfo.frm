VERSION 5.00
Begin VB.Form InventoryInfo 
   BackColor       =   &H80000001&
   Caption         =   "Inventory Reports"
   ClientHeight    =   3090
   ClientLeft      =   5535
   ClientTop       =   4905
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack2 
      BackColor       =   &H000040C0&
      Caption         =   "Go Back To Status"
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9840
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00808000&
      Caption         =   ".........."
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000040C0&
      Caption         =   "Return To POS"
      Height          =   495
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9840
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000040C0&
      Caption         =   "Exit Program"
      Height          =   495
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9840
      Width           =   2415
   End
   Begin VB.CommandButton SearchGreater 
      BackColor       =   &H00008000&
      Caption         =   "Search For Items Greater Than A Given Price"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8640
      Width           =   2415
   End
   Begin VB.CommandButton cmdSearchLess 
      BackColor       =   &H00008000&
      Caption         =   "Search For Items Less Than A Given Price"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdSortExpensive 
      BackColor       =   &H00C000C0&
      Caption         =   "$$$ - $"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortCheap 
      BackColor       =   &H00C000C0&
      Caption         =   "$ - $$$"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picSortResults2 
      BackColor       =   &H80000005&
      Height          =   9375
      Left            =   9360
      ScaleHeight     =   9315
      ScaleWidth      =   5595
      TabIndex        =   3
      Top             =   240
      Width           =   5655
   End
   Begin VB.PictureBox picSortResults 
      BackColor       =   &H80000005&
      Height          =   9375
      Left            =   3360
      ScaleHeight     =   9315
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton cmdSortZA 
      BackColor       =   &H00C000C0&
      Caption         =   "Z - A"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortAZ 
      BackColor       =   &H00C000C0&
      Caption         =   "A - Z"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblThree 
      BackColor       =   &H0000FFFF&
      Caption         =   "Search Inventory By Price"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Label lblTwo 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort Inventory By Price"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lblOne 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort Inventory By Name"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "InventoryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
        'jump between forms
        InventoryInfo.Visible = False
        PointOfSale.Visible = True
End Sub

Private Sub cmdBack2_Click()
        'jump between forms
        InventoryInfo.Visible = False
        Status.Visible = True
        
End Sub

Private Sub cmdClear_Click()
        'clear picture boxes
        picSortResults.Cls
        picSortResults2.Cls
End Sub

Private Sub cmdQuit_Click()
        'quit program
        End
End Sub

Private Sub cmdSearchLess_Click()
        'search for items less than a given price
        Dim RSearch As Single
                       
        Open App.Path & "\RetailPOSAndInventoryControl.txt" For Input As #1
        Pos = 0
        Do Until EOF(1)
            Input #1, SItems, SPrices, SStartInv
            Pos = Pos + 1
            Items(Pos) = SItems
            Prices(Pos) = SPrices
            StartInv(Pos) = SStartInv
        Loop
        Close #1
        
        RSearch = InputBox("Input A Search Price", "Search Prices")
        
        Counter = 40
        Pos = 0
    
        picSortResults2.Cls
        picSortResults2.Print "  "
        picSortResults2.Print Tab(5); "Items"; Tab(35); "Prices"
        picSortResults2.Print "  "
        
        Do While (Pos < Counter)
            Pos = Pos + 1
                If Prices(Pos) < RSearch Then
                    picSortResults2.Print Tab(2); Items(Pos); Tab(35); Prices(Pos)
                End If
        Loop
        
End Sub

Private Sub cmdSortAZ_Click()
        'sort item names alphabetically going from a to z
        picSortResults.Cls
        picSortResults.Print "  "
        picSortResults.Print Tab(5); "Items"; Tab(35); "Prices"; Tab(45); "Starting Inventory"
        picSortResults.Print "  "
        
        'Load info from file
        Pos = 1
        Open App.Path & "\RetailPOSandInventoryControl.txt" For Input As #1
        Do Until EOF(1)
            Input #1, SItems, SPrices, SStartInv
            Pos = Pos + 1
            Size = Size + 1
            Items(Pos) = SItems
            Prices(Pos) = SPrices
            StartInv(Pos) = SStartInv
        Loop
        Close #1
        'Save array size
        
        Size = Pos
        
        For Pass = 1 To Size
            For Pos = 1 To (Size - Pass)
                If Items(Pos) < Items(Pos + 1) Then
                    TempItems = Items(Pos)
                    Items(Pos) = Items(Pos + 1)
                    Items(Pos + 1) = TempItems
                    
                    'also swap corresponding name in parallel arrays
                    
                    TempPrices = Prices(Pos)
                    Prices(Pos) = Prices(Pos + 1)
                    Prices(Pos + 1) = TempPrices
                    TempStartInv = StartInv(Pos)
                    StartInv(Pos) = StartInv(Pos + 1)
                    StartInv(Pos + 1) = TempStartInv
                End If
            Next Pos
            picSortResults.Print Tab(2); Items(Pos); Tab(35); Prices(Pos); Tab(50); StartInv(Pos)
        Next Pass
                                   
End Sub

Private Sub cmdSortExpensive_Click()
        'sort items by price going from most to least expensive
        picSortResults.Cls
        picSortResults.Print "  "
        picSortResults.Print Tab(5); "Items"; Tab(35); "Prices"; Tab(45); "Starting Inventory"
        picSortResults.Print "  "
                       
        For Pass = 1 To Size
            For Pos = 1 To (Size - Pass)
                If Prices(Pos) < Prices(Pos + 1) Then
                    TempPrices = Prices(Pos)
                    Prices(Pos) = Prices(Pos + 1)
                    Prices(Pos + 1) = TempPrices
                    
                    'also swap corresponding name in parallel arrays
                    
                    TempItems = Items(Pos)
                    Items(Pos) = Items(Pos + 1)
                    Items(Pos + 1) = TempItems
                    TempStartInv = StartInv(Pos)
                    StartInv(Pos) = StartInv(Pos + 1)
                    StartInv(Pos + 1) = TempStartInv
                End If
            Next Pos
            picSortResults.Print Tab(2); Items(Pos); Tab(35); Prices(Pos); Tab(50); StartInv(Pos)
        Next Pass
End Sub

Private Sub cmdSortZA_Click()
        'sort items by name reverse alphabetically going from z to a
        picSortResults.Cls
        picSortResults.Print "  "
        picSortResults.Print Tab(5); "Items"; Tab(35); "Prices"; Tab(45); "Starting Inventory"
        picSortResults.Print "  "
        
         
        
        For Pass = 1 To Size
            For Pos = 1 To (Size - Pass)
                If Items(Pos) > Items(Pos + 1) Then
                    TempItems = Items(Pos)
                    Items(Pos) = Items(Pos + 1)
                    Items(Pos + 1) = TempItems
                    
                    'also swap corresponding name in parallel arrays
                    
                    TempPrices = Prices(Pos)
                    Prices(Pos) = Prices(Pos + 1)
                    Prices(Pos + 1) = TempPrices
                    TempStartInv = StartInv(Pos)
                    StartInv(Pos) = StartInv(Pos + 1)
                    StartInv(Pos + 1) = TempStartInv
                End If
            Next Pos
            picSortResults.Print Tab(2); Items(Pos); Tab(35); Prices(Pos); Tab(50); StartInv(Pos)
        Next Pass
End Sub

Private Sub cmdSortCheap_Click()
        'sort items by price going from the cheapest to the most expensive
        picSortResults.Cls
        picSortResults.Print "  "
        picSortResults.Print Tab(5); "Items"; Tab(35); "Prices"; Tab(45); "Starting Inventory"
        picSortResults.Print "  "
                       
        For Pass = 1 To Size
            For Pos = 1 To (Size - Pass)
                If Prices(Pos) > Prices(Pos + 1) Then
                    TempPrices = Prices(Pos)
                    Prices(Pos) = Prices(Pos + 1)
                    Prices(Pos + 1) = TempPrices
                    
                    'also swap corresponding name in parallel arrays
                    
                    TempItems = Items(Pos)
                    Items(Pos) = Items(Pos + 1)
                    Items(Pos + 1) = TempItems
                    TempStartInv = StartInv(Pos)
                    StartInv(Pos) = StartInv(Pos + 1)
                    StartInv(Pos + 1) = TempStartInv
                End If
            Next Pos
            picSortResults.Print Tab(2); Items(Pos); Tab(35); Prices(Pos); Tab(50); StartInv(Pos)
        Next Pass
                     
        

End Sub

Private Sub SearchGreater_Click()
        'search for items with prices greater than that of an input price
        Dim RSearch As Single
                       
        Open App.Path & "\RetailPOSAndInventoryControl.txt" For Input As #1
        Pos = 0
        Do Until EOF(1)
            Input #1, SItems, SPrices, SStartInv
            Pos = Pos + 1
            Items(Pos) = SItems
            Prices(Pos) = SPrices
            StartInv(Pos) = SStartInv
        Loop
        Close #1
        
        RSearch = InputBox("Input A Search Price", "Search Prices")
        
        Counter = 40
        Pos = 0
    
        picSortResults2.Cls
        picSortResults2.Print "  "
        picSortResults2.Print Tab(5); "Items"; Tab(35); "Prices"
        picSortResults2.Print "  "
        
        Do While (Pos < Counter)
            Pos = Pos + 1
                If Prices(Pos) > RSearch Then
                    picSortResults2.Print Tab(2); Items(Pos); Tab(35); Prices(Pos)
                End If
        Loop
        
End Sub



'RetailPOSandInventoryControl program; InventoryInfo form
'this code was written on Thursday, November 2, 2006
'written by Mark Collette
'the purpose of this form is to organize items in several different manners
'the subroutines sorted data based on item name, and price
'the subroutines also searched item prices and displayed items with prices greater than and/or less than of the input price
