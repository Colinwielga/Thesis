VERSION 5.00
Begin VB.Form frm2Menu 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   ScaleHeight     =   629
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   917
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFinishedOrder 
      Caption         =   "Done Ordering"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2280
      TabIndex        =   17
      Top             =   7680
      Width           =   2775
   End
   Begin VB.PictureBox PicDeepDish 
      Height          =   3735
      Left            =   7920
      Picture         =   "frmMenu2.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   4635
      TabIndex        =   16
      Top             =   5400
      Width           =   4695
   End
   Begin VB.CommandButton cmdRefreshCart 
      Caption         =   "Show Cart"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   12000
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalad 
      Caption         =   "Caeser Salad"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   12
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdDeluxeGB 
      Caption         =   "Deluxe Garlic Bread"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   8
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdGarlicBread 
      Caption         =   "Garlic Bread"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdCalzones 
      Caption         =   "Calzones"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton cmdBarBQChix 
      Caption         =   "Bar-B-Q Chicken"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton cmdChickenRoma 
      Caption         =   "Chicken and Roma Tomato"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdTommySpecial 
      Caption         =   "Tommy Special"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdStuffedPie 
      Caption         =   "Italian Stuffed Pie"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton cmdDeepDish 
      Caption         =   "Chicago Stlye Deep Dish"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   2415
   End
   Begin VB.PictureBox PicCart 
      Height          =   3015
      Left            =   7920
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   14
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblSpecialtyPizzas 
      BackColor       =   &H000000C0&
      Caption         =   "Specialty Pizzas"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label lblApps 
      BackColor       =   &H000000C0&
      Caption         =   "Appetizers"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label lblGourmetPizzas 
      BackColor       =   &H000000C0&
      Caption         =   "Gourmet Pizzas"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label lblCart 
      BackColor       =   &H000000C0&
      Caption         =   "Cart"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frm2Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the Menu page and it gives the user a variety of different items to choose from.
'There are 9 item buttons categorized in three sections. The first section is called Specialty
'Pizzas. The second category is called Gourmet Pizzas, and the third is called Appetizers.
'There is a picture box that displays the cart items and keeps a list of what the user is ordering.

Private Sub cmdBarBQChix_Click()
    
    'Asks the user to input the number of pizzas they want
    NumberOfItem = InputBox("This kind of pizza only comes in medium size. How many pizzas would you like 1-5?", "Bar-B-Q Chicken Pizza")
    
    'This only allows the user to input number 1 through 5
    'If input a number over 5 a message box appears saying entry invalid
    'and will reask to enter in a number 1 through 5
    Do While NumberOfItem > 5
       MsgBox "Entry Invalid"
       NumberOfItem = InputBox("This kind of pizza only comes in medium size. How many pizzas would you like 1-5?", "Bar-B-Q Chicken Pizza")
    Loop
     
    
  'load our cart array
  'this for loop will automatically place the chosen item and how many in the cart list
    Dim i As Integer
    For i = 1 To NumberOfItem
        'Add number to the cart list
        incramentCartList
        pizzaList(pizzaListCtr, 0) = "Bar-B-Q Chicken"
    Next i
        
        
    
    
End Sub

Private Sub cmdCalzones_Click()
    'This is the variable to change the title of the third form to Calzones
    pizzaChoiceTitle = "Calzones"
    
    'Add number to cart list
    incramentCartList
    
    'Add own toppings to the calzone if the user chooses to
    frm3Toppings.addDataToComboBox
    
    'Switches the form from Menu to Toppings and displays Calzones as the title on the toppings form
    frm2Menu.Hide
    frm3Toppings.Show
    frm3Toppings.ChangeTitle ("Calzones")
    
    'Clear pic in toppings form
    frm3Toppings.picToppingResults.Cls
    
    
End Sub

Private Sub cmdChickenRoma_Click()
    
    'Asks the user to input the number of pizzas they want
    NumberOfItem = InputBox("This kind of pizza only comes in medium size. How many pizzas would you like 1-5?", "Chicken and Roma Tomato.")
    
    'This only allows the user to input number 1 through 5
    'If input a number over 5 a message box appears saying entry invalid
    'and will reask to enter in a number 1 through 5
    Do While NumberOfItem > 5
       MsgBox "Entry Invalid"
       NumberOfItem = InputBox("This kind of pizza only comes in medium size. How many pizzas would you like 1-5?", "Chicken and Roma Tomato.")
    Loop
    
    
    'display the order in the picture box
    'PicCart.Print NumberOfItem; " "; Size; " Chicken and Roma"
   
    'load our cart array
    'this for loop will automatically place the chosen item and how many in the cart list
    Dim i As Integer
    For i = 1 To NumberOfItem
        'Add number to the cart list
        incramentCartList
        pizzaList(pizzaListCtr, 0) = "Chicken and Roma Tomato"
    Next i
   
End Sub

Private Sub cmdDeepDish_Click()
    'declare the variables
    Dim Price As Single
    
    'This is the variable to change the title of the third form to Chicago Style Deep Dish
    pizzaChoiceTitle = "Chicago Style Deep Dish"
    
    'Add number to the cart list
    incramentCartList
    
    'Switches the form from Menu to Toppings and displays Chicago Style Deep Dish as the title on the toppings form
    frm2Menu.Hide
    frm3Toppings.Show
    frm3Toppings.ChangeTitle ("Chicago Style Deep Dish")
    'Add own toppings to the Chicago Deep Dish pizza if the user chooses to
    frm3Toppings.addDataToComboBox

    
    
End Sub

Private Sub cmdDeluxeGB_Click()
       'This is just given a value of 1
    'depending on how many times the item is chosen it will show up multiple times
    'in the cart list but always have a 1 before the title "Deluxe Garlic Bread"
    NumberOfItem = 1
    
    'display the order in the picture box
   ' PicCart.Print NumberOfItem; " "; Size; " Deluxe Garlic Bread "
    
     'load our cart array
      'this for loop will automatically place the chosen item and how many in the cart list
    Dim i As Integer
    For i = 1 To NumberOfItem
        'Add number to the cart list
        incramentCartList
        pizzaList(pizzaListCtr, 0) = "Deluxe Garlic Bread"
    Next i
   

End Sub

Private Sub cmdFinishedOrder_Click()
    'Switches from form 2 to 5
    frm2Menu.Hide
    frm5Checkout.Show
End Sub

Private Sub cmdGarlicBread_Click()
   
    'This is just given a value of 1
    'depending on how many times the item is chosen it will show up multiple times
    'in the cart list but always have a 1 before the title " Garlic Bread "
    NumberOfItem = 1
    
    'load our cart array
    'this for loop will automatically place the chosen item and how many in the cart list
    Dim i As Integer
    For i = 1 To NumberOfItem
        'Add number to the cart list
        incramentCartList
        pizzaList(pizzaListCtr, 0) = "Garlic Bread"
    Next i
  
End Sub

Private Sub cmdRefreshCart_Click()
    'each time the program is run the cls function clears the table for it to be run again
    PicCart.Cls
               
    'loop through our cart array to display contents to customer
    Dim i As Integer
    For i = 1 To pizzaListCtr
        PicCart.Print pizzaList(i, 0)
    Next i
    
    
    
End Sub

Private Sub cmdSalad_Click()
    'This is just given a value of 1
    'depending on how many times the item is chosen it will show up multiple times
    'in the cart list but always have a 1 before the title "Caesar Salad"
    
    NumberOfItem = 1
    
    'load our cart array
    'this for loop will automatically place the chosen item and how many in the cart list
    Dim i As Integer
    For i = 1 To NumberOfItem
        'Add number to the cart list
        incramentCartList
        pizzaList(pizzaListCtr, 0) = "Caesar Salad"
    Next i
   
End Sub

Private Sub cmdStuffedPie_Click()
    'Add number to the cart list
    incramentCartList
    
    'This is the variable to change the title of the third form to Italian Stuffed Pie
    pizzaChoiceTitle = "Italian Stuffed Pie"
    
    'Add own toppings to the Italian Stuffed Pie if the user chooses to
    frm3Toppings.addDataToComboBox
    
    'Switches the form from Menu to Toppings and displays the title as Italian Stuffed Pie
    frm2Menu.Hide
    frm3Toppings.Show
    frm3Toppings.ChangeTitle ("Italian Stuffed Pie")
    
    'Clear pic in toppings form
    frm3Toppings.picToppingResults.Cls
    
    
End Sub

Private Sub cmdTommySpecial_Click()
    
    'Asks the user to input the number of pizzas they want
    NumberOfItem = InputBox("This kind of pizza only comes in medium size. How many pizzas would you like 1-5?", "Tommy Special")
    
    'This only allows the user to input number 1 through 5
    'If input a number over 5 a message box appears saying entry invalid
    'and will reask to enter in a number 1 through 5
    Do While NumberOfItem > 5
       MsgBox "Entry Invalid"
       NumberOfItem = InputBox("This kind of pizza only comes in medium size. How many pizzas would you like 1-5?", "Tommy Special")
     Loop
     
     'load our cart array
     'this for loop will automatically place the chosen item and how many in the cart list
    Dim i As Integer
    For i = 1 To NumberOfItem
        'Add number to the cart list
        incramentCartList
        pizzaList(pizzaListCtr, 0) = "Tommy Special"
    Next i
  
    
End Sub

Private Sub incramentCartList()
     'Incraments the pizza choices the user chooses in the cart
     pizzaListCtr = pizzaListCtr + 1
     
     


End Sub




