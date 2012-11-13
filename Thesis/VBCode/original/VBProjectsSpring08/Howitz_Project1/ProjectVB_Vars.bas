Attribute VB_Name = "Module1"
'Project Name: Sexton Cash Register
'Form Name:  Module
'Louis Howitz
'March 31, 2008
'These are all of the arrays in text files.  They are used in  multiple
'forms in order to keep a running total available in the frmPay.from.

Option Explicit

Public GrillFood(1 To 11) As String
Public GrillPrice(1 To 11) As Single
Public BakeFood(1 To 6) As String
Public BakePrice(1 To 6) As Single
Public Drink(1 To 9) As String
Public DrinkPrice(1 To 9) As Single
Public Deli(1 To 9) As String
Public DeliPrice(1 To 9) As Single
Public Pizza(1 To 6) As String
Public PizzaPrice(1 To 6) As Single
Public Salad(1 To 6) As String
Public SaladPrice(1 To 6) As Single
Public Snack(1 To 9) As String
Public SnackPrice(1 To 9) As Single
Public Soup(1 To 4) As String
Public SoupPrice(1 To 4) As Single


Public ShoppingCart(1 To 500) As String
Public CartPrices(1 To 500) As Single
Public Items As Integer





