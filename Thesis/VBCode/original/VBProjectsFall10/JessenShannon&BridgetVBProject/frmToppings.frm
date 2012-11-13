VERSION 5.00
Begin VB.Form frm3Toppings 
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
   Begin VB.ComboBox ComboQuantity 
      Height          =   315
      ItemData        =   "frmToppings.frx":0000
      Left            =   10800
      List            =   "frmToppings.frx":0010
      TabIndex        =   16
      Text            =   "How Many?"
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddToToppings 
      Caption         =   "Add"
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
      Left            =   8640
      TabIndex        =   15
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox txtChooseYourOwn 
      Height          =   615
      Left            =   8640
      TabIndex        =   13
      Top             =   3120
      Width           =   4335
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done With Toppings"
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
      Left            =   4440
      TabIndex        =   12
      Top             =   7800
      Width           =   2775
   End
   Begin VB.PictureBox picToppingResults 
      Height          =   4215
      Left            =   4440
      ScaleHeight     =   4155
      ScaleWidth      =   2715
      TabIndex        =   10
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdChicken 
      Caption         =   "Chicken"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdPepperoni 
      Caption         =   "Pepperoni"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdMushrooms 
      Caption         =   "Mushrooms"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdGreenPeppers 
      Caption         =   "Green Peppers"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmditaliansausage 
      Caption         =   "Italian Sausage"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdCanadianBacon 
      Caption         =   "Canadian Bacon"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPineapple 
      Caption         =   "Pineapple"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdOnions 
      Caption         =   "Onions"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdBlackOlives 
      Caption         =   "Black Olives"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblEnterYourOwn 
      BackColor       =   &H000000C0&
      Caption         =   "Enter Your Own Topping"
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
      Left            =   8640
      TabIndex        =   14
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label lblDirections 
      BackColor       =   &H000000C0&
      Caption         =   "Please Select your toppings."
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
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label lblToppingsList 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   11055
   End
End
Attribute VB_Name = "frm3Toppings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub onLoad()
    'Tells that the title will change when a button is clicked in form 2
    lblToppingsList.Caption = pizzaChoiceTitle

End Sub

    
Public Sub ChangeTitle(title)
    'Switches the title to whatever pizza the user chose in form 2
    'by useing a public sub that we gave ChangeTitle click (title) to
    'the function to do this is take the label box. caption and equal it
    'to the title and that will represent whatever pizza they chose
    'this only corresponds to the three specialty pizzas because they
    'only get to choose toppings
    
    lblToppingsList.Caption = title
    
End Sub

Public Sub addDataToComboBox()
  
  'see tutorial @ http://www.vb6.us/tutorials/visual-basic-combo-box-tutorial
  ComboQuantity.Clear
  
  'declare variable
  Dim i As Integer
 
  'This is when the user adds their own toppings and chooses how many they want.
  'This for loop will allow a user to input 1 to 10 toppings
  For i = 1 To 10
   ComboQuantity.AddItem Conversion.Str(i)
  Next i
    
  'Starts the quantity in the combo list at 0
  ComboQuantity.ListIndex = 0
   

End Sub

Private Sub cmdAddToToppings_Click()
    'This is called when someone tries to add their own topping
    'The user can input any item they want into the text box
    'If they do not enter anything and click add a msgbox with appear and say entry invalid
    
    Dim theCustomTopping As String
    If txtChooseYourOwn.Text <> "" Then
        theCustomTopping = txtChooseYourOwn.Text
        Else
        MsgBox " Entry Invalid "
    End If
    
    'This will add the number that they inputed to the toppings display box
    'The combo box only allows up to a quantity of 10 of the items inputed
    Dim i As Integer
    For i = 1 To ComboQuantity.ListIndex + 1
        picToppingResults.Print theCustomTopping
    Next i
    

End Sub

'The toppings form has 9 different choices of toppings to choose from.
'which ever topping is chosen will print in the toppings box

Private Sub cmdBlackOlives_Click()
    lblToppingsList.Caption = pizzaChoiceTitle
    picToppingResults.Print " Black Olives "

'the temp variable is specific for the pizza chosen and will reset once another pizza is chosen
'This code explains only for the different toppings
    tempI = tempI + 1
    tempToppingsList(tempI) = " Black Olives "
 
End Sub

Private Sub cmdCanadianBacon_Click()
    lblToppingsList.Caption = pizzaChoiceTitle
    picToppingResults.Print " Canadian Bacon "

'the temp variable is specific for the pizza chosen and will reset once another pizza is chosen
'This code explains only for the different toppings
    tempI = tempI + 1
    tempToppingsList(tempI) = " Canadian Bacon "
    
End Sub

Private Sub cmdChicken_Click()
    lblToppingsList.Caption = pizzaChoiceTitle
    picToppingResults.Print " Chicken "

'the temp variable is specific for the pizza chosen and will reset once another pizza is chosen
'This code explains only for the different toppings
    tempI = tempI + 1
    tempToppingsList(tempI) = " Chicken "
    
End Sub

Private Sub cmdGreenPeppers_Click()
    lblToppingsList.Caption = pizzaChoiceTitle
    picToppingResults.Print " Green Peppers "

'the temp variable is specific for the pizza chosen and will reset once another pizza is chosen
'This code explains only for the different toppings
    tempI = tempI + 1
    tempToppingsList(tempI) = " Green Peppers "
End Sub

Private Sub cmditaliansausage_Click()
    lblToppingsList.Caption = pizzaChoiceTitle
    picToppingResults.Print "Italian Sausage "

'the temp variable is specific for the pizza chosen and will reset once another pizza is chosen
'This code explains only for the different toppings
    tempI = tempI + 1
    tempToppingsList(tempI) = "Italian Sausage"
    
End Sub


Private Sub cmdMushrooms_Click()
    lblToppingsList.Caption = pizzaChoiceTitle
    picToppingResults.Print " Mushrooms "

'the temp variable is specific for the pizza chosen and will reset once another pizza is chosen
'This code explains only for the different toppings
    tempI = tempI + 1
    tempToppingsList(tempI) = " Mushrooms "
    
End Sub

Private Sub cmdOnions_Click()
    lblToppingsList.Caption = pizzaChoiceTitle
    picToppingResults.Print " Onions "

'the temp variable is specific for the pizza chosen and will reset once another pizza is chosen
'This code explains only for the different toppings
    tempI = tempI + 1
    tempToppingsList(tempI) = " Onions "
    
End Sub

Private Sub cmdPepperoni_Click()
    lblToppingsList.Caption = pizzaChoiceTitle
    picToppingResults.Print " Pepperoni "

'the temp variable is specific for the pizza chosen and will reset once another pizza is chosen
'This code explains only for the different toppings
    tempI = tempI + 1
    tempToppingsList(tempI) = " Pepperonni "
    
End Sub

Private Sub cmdPineapple_Click()
    lblToppingsList.Caption = pizzaChoiceTitle
    picToppingResults.Print " Pineapple "

'the temp variable is specific for the pizza chosen and will reset once another pizza is chosen
'This code explains only for the different toppings
    tempI = tempI + 1
    tempToppingsList(tempI) = " Pineapple "
    
    
End Sub
Private Sub cmdDone_Click()
 
    'In the toppings list when the user clicks done it will take all the items chosen
    'which is the pizzaList array and save this information in a format
    'that first names all the titles of the menu items selected and save it in column 0
    'which is assigned the pizzaList in the array
    
    pizzaList(pizzaListCtr, 0) = pizzaChoiceTitle
    'This loop is made for the array
    'only specific to the three pizzas that take toppings
    
    Dim i As Integer
    For i = 1 To 48
        pizzaList(pizzaListCtr, i) = tempToppingsList(i)
    Next i
    
    'Clear pic in toppings form
    picToppingResults.Cls
    tempI = 0
    For i = 1 To 48
        tempToppingsList(i) = ""
    Next i
    
    'switches from form 3 to 2
    frm3Toppings.Hide
    frm2Menu.Show
    
End Sub
Private Sub cmdGoToCheckout_Click()
    'Switches from form 3 to form 5
    frm3Toppings.Hide
    frm5Checkout.Show
    
End Sub

Private Sub cmdReturnMenu_Click()
    'Switches from form 2 to 3
    frm2Menu.Show
    frm3Toppings.Hide
End Sub


