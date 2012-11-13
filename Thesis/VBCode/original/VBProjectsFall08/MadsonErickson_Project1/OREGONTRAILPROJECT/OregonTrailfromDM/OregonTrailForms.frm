VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form3"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4335
      Left            =   2280
      ScaleHeight     =   4275
      ScaleWidth      =   5115
      TabIndex        =   18
      Top             =   8280
      Width           =   5175
   End
   Begin VB.CommandButton cmdSee 
      Caption         =   "See Items Normally for Sale in the Video Game "
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
      Left            =   360
      TabIndex        =   17
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton cmdAlphabetize 
      Caption         =   "Alphabetize Items -(If you're into that sorta thing)"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      TabIndex        =   16
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdRHome 
      Caption         =   "Return to Home Page"
      Height          =   1935
      Left            =   8040
      TabIndex        =   15
      Top             =   2280
      Width           =   3255
   End
   Begin VB.PictureBox picResults1 
      Height          =   3015
      Left            =   3480
      ScaleHeight     =   2955
      ScaleWidth      =   4395
      TabIndex        =   12
      Top             =   4680
      Width           =   4455
   End
   Begin VB.TextBox txtQuantity 
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox txtItem1 
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Quantity of Desired Item"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label lblDesire 
      Caption         =   "Enter Desired Item (1-8)"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblLuxury 
      Caption         =   "8. Playing Cards ($.50)"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "7. A Whip ($5)"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label lblTongue 
      Caption         =   "6. Spare Axel  and Tongue ($15)"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label lblFood 
      Caption         =   "5. 50lbs of Food ($10)"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label lblJohnWayne 
      Caption         =   "4. John Wayne Plastic Figurine ($40)"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblBullets 
      Caption         =   "3. Bullets (20 per box) ($5)"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblClothes 
      Caption         =   "2. Pair of Clothes ($5)"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblOxen 
      Caption         =   "1. Oxen ($20)"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblCost 
      Caption         =   $"Oregon Trail Forms.frx":0000
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
      Top             =   960
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
      Top             =   240
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


Dim CTR As Integer
Dim Item(1 To 100) As String
Dim Cost(1 To 100) As Single


Private Sub cmdAlphabetize_Click()

Dim Temp As String
Dim TempCost As Single
Dim TempCTR As Integer
Dim Pass As Integer
Dim Pos As Integer
Dim I As Integer


    picResults.Cls
    
        picResults.Print "Item"; Tab(28); "Cost"
            picResults.Print "======================================================="

    For Pass = 1 To CTR - 1 'keep track of how many passes
        For Pos = 1 To CTR - Pass 'keep track of how many comparisons
            If Item(Pos) > Item(Pos + 1) Then
            Temp = Item(Pos)
            Item(Pos) = Item(CTR)(Pos + 1)
            Item(Pos + 1) = Temp
            
            TempCost = Cost(CTR)(Pos)
            Cost(Pos) = Cost(CTR)(Pos + 1)
            Cost(Pos + 1) = TempCost
            
            
        End If
    Next Pos
    Next Pass
    
    'print the sorted list
    For I = 1 To CTR
         picResults.Print "Item"; Tab(28); "Cost"
       
    Next I
    


End Sub

Private Sub cmdRHome_Click()

    Form3.Hide
    Form2.Show
    

End Sub

Private Sub cmdSee_Click() 'this section loads the data about the store for the viewer to see

    picResults.Cls
     CTR = 0
    
    Open App.Path & "\GeneralStoreItems.txt" For Input As #1 'read file
    
    
     picResults.Print "Item"; Tab(28); "Cost"
            picResults.Print "======================================================="
        Do Until EOF(1)
        CTR = CTR + 1
            Input #1, Item(CTR), Cost(CTR)
            
            picResults.Print (CTR); Tab(5); Item(CTR); Tab(28); FormatCurrency(Cost(CTR))
            
      
            Loop
                Close #1
                
           picResults.Print "======================================================="
        
    
        


End Sub

Private Sub Text2_Change()

End Sub

