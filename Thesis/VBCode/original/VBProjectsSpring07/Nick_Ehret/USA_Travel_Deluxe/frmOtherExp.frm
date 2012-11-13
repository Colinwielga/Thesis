VERSION 5.00
Begin VB.Form frmOtherExp 
   BackColor       =   &H000000C0&
   Caption         =   "Other Expenses"
   ClientHeight    =   8565
   ClientLeft      =   885
   ClientTop       =   870
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   10350
   Begin VB.PictureBox picMisc 
      Height          =   1215
      Left            =   8520
      Picture         =   "frmOtherExp.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   18
      Top             =   4920
      Width           =   1575
   End
   Begin VB.PictureBox picGift 
      Height          =   1215
      Left            =   8520
      Picture         =   "frmOtherExp.frx":0917
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   17
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox picCity 
      Height          =   1215
      Left            =   8520
      Picture         =   "frmOtherExp.frx":183A
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   16
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox picFood 
      Height          =   1215
      Left            =   8520
      Picture         =   "frmOtherExp.frx":208A
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Compute the Total Cost of the Trip"
      Height          =   975
      Left            =   3000
      TabIndex        =   14
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   7800
      TabIndex        =   13
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   975
      Left            =   5400
      TabIndex        =   12
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdTotalOther 
      Caption         =   "Compute Total Other Trip Expenses"
      Height          =   975
      Left            =   600
      TabIndex        =   11
      Top             =   7440
      Width           =   2055
   End
   Begin VB.PictureBox picResultsTotal 
      Height          =   855
      Left            =   3600
      ScaleHeight     =   795
      ScaleWidth      =   6435
      TabIndex        =   10
      Top             =   6360
      Width           =   6495
   End
   Begin VB.PictureBox picResultsMisc 
      Height          =   975
      Left            =   2760
      ScaleHeight     =   915
      ScaleWidth      =   5475
      TabIndex        =   8
      Top             =   5040
      Width           =   5535
   End
   Begin VB.PictureBox picResultsGift 
      Height          =   975
      Left            =   2760
      ScaleHeight     =   915
      ScaleWidth      =   5475
      TabIndex        =   7
      Top             =   3720
      Width           =   5535
   End
   Begin VB.PictureBox picResultsSite 
      Height          =   975
      Left            =   2760
      ScaleHeight     =   915
      ScaleWidth      =   5475
      TabIndex        =   6
      Top             =   2400
      Width           =   5535
   End
   Begin VB.PictureBox picResultsFood 
      Height          =   975
      Left            =   2760
      ScaleHeight     =   915
      ScaleWidth      =   5475
      TabIndex        =   5
      Top             =   1080
      Width           =   5535
   End
   Begin VB.CommandButton cmdMisc 
      Caption         =   "Miscellaneous Expenses"
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdGifts 
      Caption         =   "Gift Expense"
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdSiteSeeing 
      Caption         =   "Site Seeing Expense"
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdFood 
      Caption         =   "Food Expense"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblTotalOther 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Total Other Trip Expenses"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Label lblOtherExp 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Other Trip Expenses"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmOtherExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: USA Travel Planner Deluxe
'Written By: Nick Ehret
'Date Written: March 21, 2007
'Form Purpose: 'The purpose of this form is to calculate some other expenses of the trip
                'for the user. It will calculate the food, site seeing, gift, and miscellaneous
                'expenses of the trip. Finally, at the end it will calculate the total cost
                'of other expenses

Option Explicit
Dim FoodExp As Single
Dim Food As Single
Dim SiteExp As Single
Dim GiftExp As Single
Dim MiscExp As Single





Private Sub cmdBack_Click()
    'This button will bring the user back to frmUSATravel
    
    frmTravel.Visible = True
    frmOtherExp.Visible = False
    
End Sub

Private Sub cmdExit_Click()
    'This button will end the program
    End
End Sub

Private Sub cmdFood_Click()
    'this button will ask the traveler how much they want to spend on food a day during
    'their trip
    
    picResultsFood.Cls
    
    Food = InputBox("How much do you prefer to spend on food during a day on your trip", "Food Expense")
    FoodExp = Food * 7 'calculates the cost of food based on a weeklong stay
    
    picResultsFood.Print "Your preferred cost for food on your trip is "; FormatCurrency(FoodExp)
    
    
End Sub

Private Sub cmdGifts_Click()
    'this button will ask the traveler how much they want to spend on gifts during
    'their trip
    
    picResultsGift.Cls
    
    GiftExp = InputBox("How much do you prefer to spend on gifts for your family and friends during your trip", "Gift Expense")
    
    picResultsGift.Print "Your preferred cost for gifts on your trip is "; FormatCurrency(GiftExp)
    
End Sub

Private Sub cmdMisc_Click()
    'this button will ask the traveler how much they want to spend on miscellaneous
    'expenses on their trip
    
    picResultsMisc.Cls
    
    MiscExp = InputBox("How much do you prefer to spend on miscellaneous expenses during your trip", "Miscellaneous Expenses")
    
    picResultsMisc.Print "Your preferred cost for miscellaneous expenses on your trip is "; FormatCurrency(MiscExp)
    
End Sub

Private Sub cmdSiteSeeing_Click()
    'this button will ask the traveler how much they want to spend on site seeing during
    'their trip
    
    picResultsSite.Cls
    
    SiteExp = InputBox("How much do you prefer to spend on site seeing during your trip", "Site Seeing Expense")
    
    picResultsSite.Print "Your preferred cost for site seeing on your trip is "; FormatCurrency(SiteExp)
    
End Sub

Private Sub cmdTotal_Click()
    'This button will bring the user to frmTotalCost if the user has calculated total
    'other expenses
    
    If TotalOtherExp > 0 Then
        frmTotalCost.Visible = True
        frmOtherExp.Visible = False
    Else
        MsgBox "Please calculate total other expenses", , "Error"
    End If
    
End Sub

Private Sub cmdTotalOther_Click()
    'this button will compute the total for other expenses on the travelers trip
      
    picResultsTotal.Cls
    
    TotalOtherExp = GiftExp + MiscExp + FoodExp + SiteExp 'Calculates the total cost of other expenses
    
    picResultsTotal.Print "Your total cost for other expenses on your trip is "; FormatCurrency(TotalOtherExp)
    
    
End Sub
