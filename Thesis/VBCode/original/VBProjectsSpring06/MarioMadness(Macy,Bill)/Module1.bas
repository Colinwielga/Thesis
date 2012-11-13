Attribute VB_Name = "Module1"
'Project name: Mario Madness
'Form name: Module
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This module stores the information that is needed across certain forms.
                'The login and registration is for the first chunk of variables and the second
                'section refers to the information that is shared between the magazine order
                'form and the confirmation of that information on the next form.  The purpose
                'for this module is so that those forms can work together using the same variables
                'that have been used by the user for login and the magazine order form.



Option Explicit

    Global number As Integer
    Global username(1 To 100) As String
    Global password(1 To 100) As String
    Global loginname As String
    Global loginpassword As String
    Global wrongentry As Boolean
    Global chosenname As String
    Global chosenpassword As String
    Global myvariable As Boolean
    
    Global name1 As String
    Global streetaddress As String
    Global city As String
    Global postalcode As Single
    Global emailaddress As String
    Global pobox As Single
    Global phonenumber As Integer
    Global phonenumber2 As Integer
    Global phonenumber3 As Integer
    Global cardname As String
    Global cardnumber As Double
    Global cardtype As String
    Global expdate As String
    Global expdate2 As Integer
    Global country As String
    Global state As String
    Global Date1 As String
    Global date2 As Integer
    Global paymenttype As String
    Global subscriptiontype As String
