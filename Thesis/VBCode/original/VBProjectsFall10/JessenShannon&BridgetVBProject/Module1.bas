Attribute VB_Name = "Module1"
'This is where alll our global variables will go

    'General
    Public pizzaChoiceTitle As String
    Public NumberOfItem As Integer
    Public Size As String
    
    'Customer Info
    Public FirstName As String
    Public LastName As String
    Public firstAndLast As String
    Public customerAddress As String
    
    'Cart
    Public pizzaList(0 To 50, 0 To 50) As String
    Public currentToppings(0 To 50) As String
    Public pizzaListCtr As Integer
    Public tempToppingsList(1 To 48) As String
    Public tempI As Integer
    Public colorCounter As Integer
  
    'For refernce
    '''''''''''''''''''''''''''''''
    'So at x the pizzaListCtr = 2
    'Col 0 is the names of the pizzas
    'Col 1 is the price
    
    
    'it should generally print out like this
    'deep dish price: $16
    'calzone price $8
