VERSION 5.00
Begin VB.Form frmRegister 
   Caption         =   "Register with James Delivers"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRegisterPassword 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   22
      Top             =   7200
      Width           =   4095
   End
   Begin VB.TextBox txtRegisterLogin 
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   21
      Top             =   6600
      Width           =   4095
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   495
      Left            =   480
      TabIndex        =   18
      Top             =   7800
      Width           =   2535
   End
   Begin VB.TextBox txtExpirationDate 
      Height          =   375
      Index           =   7
      Left            =   2640
      TabIndex        =   17
      Top             =   6000
      Width           =   4095
   End
   Begin VB.TextBox txtCreditCardNumber 
      Height          =   375
      Index           =   6
      Left            =   2640
      TabIndex        =   16
      Top             =   5400
      Width           =   4095
   End
   Begin VB.TextBox txtPaymentMethod 
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   15
      Top             =   4800
      Width           =   4095
   End
   Begin VB.TextBox txtZip 
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   14
      Top             =   4200
      Width           =   4095
   End
   Begin VB.TextBox txtState 
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   13
      Top             =   3600
      Width           =   4095
   End
   Begin VB.TextBox txtCity 
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   12
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   10
      Top             =   1800
      Width           =   4095
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   7800
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   120
      Picture         =   "frmRegister.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblPassword 
      Caption         =   "Choose a Password"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label lblLogin 
      Caption         =   "Select a Login"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label lblExpirationDate 
      Caption         =   "Expiration Date"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label lblCreditCardNumber 
      Caption         =   "Credit Card Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label lblPaymentMethod 
      Caption         =   "Payment Method"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblZip 
      Caption         =   "Zip Code"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblState 
      Caption         =   "State"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lblCity 
      Caption         =   "City"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblStreetAddress 
      Caption         =   "Street Address"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'James Delivers, CSCI 130 Visual Basic Project
'frmRegister
'written by James Garay Heelan
'written on 11-2-06
'This form loads the information provided by the user for registration into an array
'that will allow him or her to login into the program.  The information gathered here
'will be used later in the program for both communication and shipping purposes.

Option Explicit

Private Sub cmdLogOut_Click() 'Exits the program
    End
End Sub

Private Sub cmdSubmit_Click()

    Open App.Path & "/RegisteredUsers.txt" For Append As #1 'Opens a text file, in the same folder as the program, to be written into
        Write #1, txtName(0).Text, txtAddress(1).Text, txtCity(2).Text, txtState(3).Text, txtZip(4).Text, txtPaymentMethod(5).Text, txtCreditCardNumber(6).Text, txtExpirationDate(7).Text, txtRegisterLogin(0).Text, txtRegisterPassword(1).Text 'Writes the user's registration information into the text file for storage
    Close #1 'Closes the text file
    frmRegister.Hide 'hides this form
    frmFrontPage.Show 'brings up the login page form for the user
    
End Sub

