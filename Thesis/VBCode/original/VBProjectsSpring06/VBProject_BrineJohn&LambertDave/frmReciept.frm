VERSION 5.00
Begin VB.Form frmReciept 
   BackColor       =   &H8000000E&
   Caption         =   "Reciept"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   Picture         =   "frmReciept.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      Picture         =   "frmReciept.frx":917B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      Picture         =   "frmReciept.frx":B38E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Picture         =   "frmReciept.frx":DBFB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   5415
      Left            =   1080
      ScaleHeight     =   5355
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblStoreFront 
      BackStyle       =   0  'Transparent
      Caption         =   "Click To Go Back to Our Store Font"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   5
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label lblReciept 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click The Recipt to Print Your Final Total"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   2280
      TabIndex        =   4
      Top             =   7320
      Width           =   1935
   End
End
Attribute VB_Name = "frmReciept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dave and JB's ski and skate store
'frmFront
'Dave Lambert and John Brine
'Thursday March 23
'This form is used to print the ongoing totals that were recieved by the ski and skate store

    Dim SkateTotal2 As Double 'establises skate total #2

Private Sub cmdBack_Click()
    'goes to the store front page
    frmFront.Visible = True
    frmReciept.Visible = False

End Sub

Private Sub cmdExit_Click()
'displays thank you message and shipping notice as well as ends the program
    MsgBox "Thanks for shopping your product(s) will ship within 3 buisness days.", , "Thanks!"
    End
End Sub

Private Sub cmdShow_Click()
    'establish all necessary variables
    Dim tax, tax1 As Double
    Dim total As Double
    Dim pos As Integer
    Dim subtotal, subtotal1, skitotal As Double
    
    'writes the ski total to the file for use by site administrator
    Open App.Path & "\skitotals2.txt" For Append As #2
    Write #2, SubTotalSki
    Close #2
    'writes the skate total to the file for use by site administrator
    Open App.Path & "\skatetotals2.txt" For Append As #3
    Write #3, SubTotalSkate
    Close #3
    
    
    'first clears picture box then prints all of the user information that was input in previous page
    picResults.Cls
    picResults.Print
    picResults.Print "Name:", , Name1
    picResults.Print "Billing Address:", Address
    picResults.Print "Card Number:", , Credit
    picResults.Print "Experation Date:", Experation
    picResults.Print "Work Phone:", , Work
    picResults.Print "Home Phone:", , Home
    picResults.Print "Email:", , Email
    picResults.Print
    picResults.Print
    picResults.Print
    
    'first establishes if a user has previously selected a skate and added it to running total
    'if user has not added a skate to the running skate total then program will skip this sub routine
     If SubTotalSkate > 0 Then
            picResults.Print "--------------------------------------------Skate-------------------------------------------------"
            picResults.Print "Subtotal", , , FormatCurrency(SubTotalSkate) 'displays subtotal from running subtotal in skate store
            tax1 = SubTotalSkate * 0.07 'adds tax
            picResults.Print "Shipping and Handling", , "$25.00" 'displays shipping title and thenshipping amount
            picResults.Print "Tax", , , FormatCurrency(tax1) 'figures and displays taxes
            picResults.Print "*******************************************************************************************************************************************************"
            SkateTotal2 = 25 + subtotal1 * 1.07 'figures final amount
            picResults.Print "Total", , , FormatCurrency(SkateTotal) 'displays total
        End If
        'establishes if user has selected a ski(s)
        If SubTotalSki > 0 Then
            picResults.Print "********************************Ski**********************************************************************************"
            picResults.Print "Subtotal", , , FormatCurrency(SubTotalSki) 'displays and figures the subtotal of all selected skis
            tax = SubTotalSki * 0.07 'figures amount of tax
            picResults.Print "Shipping and Handling", , "$25.00" 'shows the amount of shipping and handling
            picResults.Print "Tax", , , FormatCurrency(tax) 'displays the amount of tax for running ski total
            picResults.Print "Ski Total", , , FormatCurrency((25 + SubTotalSki * 1.07))
            picResults.Print "-----------------------------------------------------------------------------------------------------------------------------------------------"
            picResults.Print "*******************************************************************************************************************************************************"
            picResults.Print "-----------------------------------------------------------------------------------------------------------------------------------------------"
            total = 25 + SubTotalSki * 1.07 'adds all of the different parts of the total together
            picResults.Print "Final Total", , , FormatCurrency(SkateTotal + total) 'displays the total
        End If
        
    End Sub

Private Sub picResults_Click()
    'displays error for when picture box is clicked
    MsgBox "Click the print button", , "Error!"
End Sub
