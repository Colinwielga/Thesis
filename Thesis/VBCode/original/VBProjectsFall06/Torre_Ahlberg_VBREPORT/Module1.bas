Attribute VB_Name = "Module1"
Option Explicit
'Written by Torre Ahlberg 11/02/06
'The purpose of this form is to allow the transfer of data from one form to another
'It allows the user to input their Full name in on one form, and have the program convert
'it into an ID on the next form

'This Module's purpose is to make these variables public
'So when the function on page four is looking for an input that was input on Form three
'this module allows this to happen by making the variable public

Public Name As String
Public ID As String
Public N As Integer
Public First As String, Middle As String, Last As String
Public YourName As String




