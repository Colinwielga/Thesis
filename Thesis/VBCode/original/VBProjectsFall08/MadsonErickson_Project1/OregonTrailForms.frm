VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form3"
   Picture         =   "Oregon Trail Forms.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPriceSort 
      Caption         =   "Sort by Price"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   15
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Items for Sale"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   1800
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   5400
      ScaleHeight     =   2475
      ScaleWidth      =   4035
      TabIndex        =   13
      Top             =   4800
      Width           =   4095
   End
   Begin VB.CommandButton cmdAlphabetize 
      Caption         =   "Alphabetize Items (If you're into that kind of thing)"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1680
      TabIndex        =   12
      Top             =   6240
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "           Quit Game           (and let your Oregon trail family die)"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      TabIndex        =   11
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to Shopping Cart"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5400
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdRHome 
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5400
      TabIndex        =   7
      Top             =   7560
      Width           =   4095
   End
   Begin VB.PictureBox picResults1 
      Height          =   2535
      Left            =   1680
      ScaleHeight     =   2475
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox txtQuantity 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox txtItem 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Quantity of Desired Item"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label lblDesire 
      Caption         =   "Enter Desired Item (1-8)"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lblCost 
      Caption         =   $"Oregon Trail Forms.frx":18C3A2
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   9855
   End
   Begin VB.Label lblGenStor 
      Caption         =   "Welcome to the General Store!"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This program runs the "general store". It prints data, alphabetizes data,
'and computes the cost of purchasing goods at the store. The gamer has
'$100 to spend at the store and must buy a minimum of 1 oxen, 50 Lbs of food,
'and 1 set of clothes.

Dim TempItemNumber As Integer
Dim Temp As String
Dim TempCost As Single
Dim TempCTR As Integer
Dim Pass As Integer
Dim Pos As Integer
Dim Item(1 To 100) As String
Dim ItemNumber(1 To 100) As Integer
Dim Sum As Single
Dim Cart(1 To 100) As Single

Dim CTR As Integer
Dim Cost(1 To 100) As Single


Private Sub cmdAdd_Click()
  Dim I As Integer
  Dim Q As Integer
  
        I = txtItem.Text
        Q = txtQuantity.Text
        
        CTR = 1
        
        Do While I <> ItemNumber(CTR)
        CTR = CTR + 1
        
        Loop
        Sum = Sum + (Q * Cost(CTR))

    picResults.Print FormatCurrency((Q) * (Cost(CTR))); Tab(15); Item(CTR)
    picResults.Print "======================================================="



End Sub

Private Sub cmdAlphabetize_Click()

    picResults1.Cls
    Dim I As Integer
    
       picResults1.Print "Item#"; Tab(9); "Item"; Tab(28); "Cost"
            picResults1.Print "======================================================="

    For Pass = 1 To CTR - 1 'keep track of how many passes
        For Pos = 1 To CTR - Pass 'keep track of how many comparisons
            If Item(Pos) > Item(Pos + 1) Then
            Temp = Item(Pos)
            Item(Pos) = Item(Pos + 1)
            Item(Pos + 1) = Temp
            
            TempCost = Cost(Pos)
            Cost(Pos) = Cost(Pos + 1)
            Cost(Pos + 1) = TempCost
            
            TempItemNumber = ItemNumber(Pos)
            ItemNumber(Pos) = ItemNumber(Pos + 1)
            ItemNumber(Pos + 1) = TempItemNumber
            
            
        End If
    Next Pos
    Next Pass
    
    'print the sorted list
    For I = 1 To CTR
        picResults1.Print ItemNumber(I); Tab(9); Item(I); Tab(32); FormatCurrency(Cost(I))
       
    Next I
    


End Sub


Private Sub cmdPriceSort_Click()

    Dim T As Integer
    
        picResults1.Cls
        
         picResults1.Print "Item#"; Tab(9); "Item"; Tab(28); "Cost"
            picResults1.Print "======================================================="

        For Pass = 1 To CTR - 1 'keep track of how many passes
        For Pos = 1 To CTR - Pass 'keep track of how many comparisons
            If Cost(Pos) < Cost(Pos + 1) Then
            TempCost = Cost(Pos)
            Cost(Pos) = Cost(Pos + 1)
            Cost(Pos + 1) = TempCost
            
            Temp = Item(Pos)
            Item(Pos) = Item(Pos + 1)
            Item(Pos + 1) = Temp
            
            TempItemNumber = ItemNumber(Pos)
            ItemNumber(Pos) = ItemNumber(Pos + 1)
            ItemNumber(Pos + 1) = TempItemNumber
            
            
            
        End If
    Next Pos
    Next Pass
    
    'print the sorted list
    For T = 1 To CTR
        picResults1.Print ItemNumber(T); Tab(9); Item(T); Tab(32); FormatCurrency(Cost(T))
       
    Next T

End Sub

Private Sub cmdQuit_Click()

    End

End Sub

Private Sub cmdReset_Click()

    picResults.Cls
    

End Sub

Private Sub cmdRHome_Click()

    Form3.Hide
    Form2.Show
    

End Sub


Private Sub cmdShow_Click()

    picResults1.Cls
     CTR = 0
    
    Open App.Path & "\GeneralStoreItems.txt" For Input As #1 'read file
    
    
     picResults1.Print "Item#"; Tab(9); "Item"; Tab(28); "Cost"
            picResults1.Print "======================================================="
        Do Until EOF(1)
        CTR = CTR + 1
            Input #1, ItemNumber(CTR), Item(CTR), Cost(CTR)
            
            picResults1.Print ItemNumber(CTR); Tab(5); Item(CTR); Tab(28); FormatCurrency(Cost(CTR))
            
      
            Loop
                Close #1
                

End Sub


Private Sub cmdTotal_Click()

    picResults.Print FormatCurrency(Sum); Tab(15); "Total"
     picResults.Print "======================================================="



End Sub
