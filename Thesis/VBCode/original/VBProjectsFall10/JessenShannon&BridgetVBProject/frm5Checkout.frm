VERSION 5.00
Begin VB.Form frm5Checkout 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14040
   BeginProperty Font 
      Name            =   "Goudy Old Style"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   678
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   936
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   12
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox txtlastName 
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   11
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtfirstName 
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   10
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton cmdChangecolors 
      Caption         =   "Change Backgroud colors!"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   9600
      TabIndex        =   9
      Top             =   8640
      Width           =   1335
   End
   Begin VB.CommandButton cmdSqr 
      Caption         =   "Find the Square root of your order amount"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      TabIndex        =   8
      Top             =   8640
      Width           =   1455
   End
   Begin VB.PictureBox PicCalzone 
      Height          =   4575
      Left            =   240
      Picture         =   "frm5Checkout.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   5235
      TabIndex        =   7
      Top             =   5160
      Width           =   5295
   End
   Begin VB.CommandButton cmdFinished 
      Caption         =   "Click Here When Finished"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11520
      TabIndex        =   6
      Top             =   8640
      Width           =   1575
   End
   Begin VB.PictureBox picCheckoutResults 
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   6720
      ScaleHeight     =   4515
      ScaleWidth      =   6795
      TabIndex        =   5
      Top             =   2280
      Width           =   6855
   End
   Begin VB.CommandButton cmdResult 
      Caption         =   "Calculate Totals From Price File"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   8760
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Labelmoney 
      BackColor       =   &H000000C0&
      Caption         =   "$"
      Height          =   615
      Left            =   8760
      TabIndex        =   15
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label LabelAmount 
      BackColor       =   &H000000C0&
      Height          =   495
      Left            =   9480
      TabIndex        =   14
      Top             =   7320
      Width           =   3015
   End
   Begin VB.Label LabelTotal 
      BackColor       =   &H000000C0&
      Caption         =   "Total"
      Height          =   735
      Left            =   7080
      TabIndex        =   13
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H000000C0&
      Caption         =   "Address:"
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
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label lbllastName 
      BackColor       =   &H000000C0&
      Caption         =   "Last Name: "
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
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label LabelFirstName 
      BackColor       =   &H000000C0&
      Caption         =   "First Name:"
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
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblCheckout 
      BackColor       =   &H000000C0&
      Caption         =   "Checkout"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "frm5Checkout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdChangecolors_Click()
'This button will change the first background color to the next color and there
'are 5 different colors the user can see.
'backcolor is equal to one certain color
'the format with the ampersand is the technical title of the color
'colorcounter is incremented depending on how many times you click the button

colorCounter = colorCounter + 1

If colorCounter = 1 Then
    frm5Checkout.BackColor = &HFF0000
End If
If colorCounter = 2 Then
   frm5Checkout.BackColor = &HFF00&
End If
If colorCounter = 3 Then
    frm5Checkout.BackColor = &HC000C0
End If
If colorCounter = 4 Then
    frm5Checkout.BackColor = &H80FF&
End If
If colorCounter = 5 Then
    frm5Checkout.BackColor = &HC000&
End If
If colorCounter = 6 Then
    frm5Checkout.BackColor = &HC0&
End If
    


End Sub

Private Sub cmdFinished_Click()

'This one is easy :)
End


End Sub

'Allows user to input their information into text boxes
'once they have inputed information it will print when the Results button on that page has been clicked

Private Sub cmdResult_Click()
'Declare Variables
Dim FirstName, LastName, Address As String

'clear the picture box
picCheckoutResults.Cls

'User can add in information
FirstName = txtfirstName
LastName = txtlastName
Address = txtAddress

'Prints results in picture box
picCheckoutResults.Print FirstName; " "; LastName; ", "; Address
picCheckoutResults.Print "************************************************************************************"

'Reads data from an input file
Open App.Path & "\MenuItemPrices.txt" For Input As #1

'Declare variables
Dim TotalCost As Integer
Dim Calzone_price As Integer
Dim deepDish_price, italianPie_price, bbq_price, tommys_price, chickenroma_price, garlic_price, deluxe_price, salad_price As Integer

'This declares string names for each button in the array
'Depending on the button chosen the string will equal that button name

Dim deepDish_string As String
deepDish_string = "Chicago Style Deep Dish"
Dim italianPie_string As String
italianPie_string = "Italian Stuffed Pie"
Dim bbq_string As String
bbq_string = "Bar-B-Q Chicken"
Dim tommys_string As String
tommys_string = "Tommy Special"
Dim chickenroma_string As String
chickenroma_string = "Chicken and Roma Tomato"
Dim garlic_string As String
garlic_string = "Garlic Bread"
Dim deluxe_string As String
deluxe_string = "Deluxe Garlic Bread"
Dim salad_string As String
salad_string = "Caesar Salad"
Dim Calzone_string As String
Calzone_string = "Calzones"
Dim tempPrice As String
Dim tempName As String

'This loop will read the file information
'tempPrice gets the first name in the file
'tempName gets the second name in the file
'If the tempName gets the string name that it is assigned
'then the price of that string will equal the tempPrice in the file
'This if statement is nested in a loop so that information can be filed through and read correctly

Do While Not EOF(1)
    Input #1, tempPrice, tempName
    
    If tempName = deepDish_string Then
        deepDish_price = tempPrice
    ElseIf tempName = Calzone_string Then
        Calzone_price = tempPrice
    ElseIf tempName = italianPie_string Then
        italianPie_price = tempPrice
    ElseIf tempName = bbq_string Then
        bbq_price = tempPrice
    ElseIf tempName = tommys_string Then
        tommys_price = tempPrice
    ElseIf tempName = chickenroma_string Then
        chickenroma_price = tempPrice
    ElseIf tempName = garlic_string Then
        garlic_price = tempPrice
    ElseIf tempName = deluxe_string Then
        deluxe_price = tempPrice
    ElseIf tempName = salad_string Then
        salad_price = tempPrice
    End If
        
Loop

'this for loop is taking each item that the user selected and assigning it
'to the rows and columns in the array
'this also loops through the file and calculates the total cost for each item selected
'an if statement is necessary for this for the user to choose any item they prefer
'the conversion.int converts the string function into an integer

Dim i As Integer
For i = 1 To 50
    If pizzaList(i, 0) = deepDish_string Then
        pizzaList(i, 1) = deepDish_price
        TotalCost = TotalCost + Conversion.Int(deepDish_price)
        ElseIf pizzaList(i, 0) = Calzone_string Then
        pizzaList(i, 1) = Calzone_price
        TotalCost = TotalCost + Conversion.Int(Calzone_price)
        ElseIf pizzaList(i, 0) = italianPie_string Then
        pizzaList(i, 1) = italianPie_price
        TotalCost = TotalCost + Conversion.Int(italianPie_price)
        ElseIf pizzaList(i, 0) = bbq_string Then
        pizzaList(i, 1) = bbq_price
        TotalCost = TotalCost + Conversion.Int(bbq_price)
        ElseIf pizzaList(i, 0) = tommys_string Then
        pizzaList(i, 1) = tommys_price
        TotalCost = TotalCost + Conversion.Int(tommys_price)
        ElseIf pizzaList(i, 0) = chickenroma_string Then
        pizzaList(i, 1) = chickenroma_price
        TotalCost = TotalCost + Conversion.Int(chickenroma_price)
        ElseIf pizzaList(i, 0) = garlic_string Then
        pizzaList(i, 1) = garlic_price
        TotalCost = TotalCost + Conversion.Int(garlic_price)
        ElseIf pizzaList(i, 0) = deluxe_string Then
        pizzaList(i, 1) = deluxe_price
        TotalCost = TotalCost + Conversion.Int(deluxe_price)
        ElseIf pizzaList(i, 0) = salad_string Then
        pizzaList(i, 1) = salad_price
        TotalCost = TotalCost + Conversion.Int(salad_price)
    End If
Next i
        
    'This will display in the label the total cost of the items ordered
    'the conversion.str changes a string into an integer
    
    LabelAmount.Caption = Conversion.Str(TotalCost)
    
     'loop through our cart array to display contents to customer
     Dim tempString As String
     'Dim j As Integer
    For i = 1 To pizzaListCtr
       ' tempString = tempString + Conversion.Str(i)
        tempString = tempString + ". "
        tempString = pizzaList(i, 0)
        tempString = tempString + " - Price: $"
        tempString = tempString + pizzaList(i, 1)
       
        'can't figure out why the following doesn't work
        ' tempString = tempString + " - Toppings: "
        'For j = 0 To 48
        'tempString = tempString + pizzaList(i, j)
         '   If pizzaList(i, j) = "" Then
         '   i = 48
         '   Else
         '   tempString = tempString + ", "
         '   End If
        'Next j
        
        picCheckoutResults.Print tempString
    Next i
    
    Close #1
    
    
End Sub

Private Sub cmdSqr_Click()
'Declare Variable
Dim mySqr As Double

'equation to square root the amount

mySqr = Math.Sqr(Conversion.Int(LabelAmount.Caption))

'picCheckoutResults.Print "The sqaure root of your order amount is"; mySqr
'picCheckoutResults.Print ""; mySqr
LabelAmount = mySqr

End Sub


