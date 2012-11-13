VERSION 5.00
Begin VB.Form frmDiscount 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Men"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShop 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Back to Shopping"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   2175
   End
   Begin VB.PictureBox picDiscount 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3960
      ScaleHeight     =   2955
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   2400
      Width           =   3735
   End
   Begin VB.CommandButton cmdDiscount 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Click Here to View Disounts Currently Offered"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox picBrownBag 
      Height          =   3615
      Left            =   480
      Picture         =   "frmMen.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "frmDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form will open an array and show the customer what discounts
'are currently available and print all options in a picture box
'that is easy for the customer to read.

'Declare all variables
Dim ItemsPurchased(1 To 10) As Integer
Dim PercentageDiscount(1 To 10) As Single
Dim CTR As Integer

'The disount buttonw will open the data file containing two arrays.
'It will then print the number of items purchased with its corresponding
'discount in a table on the form.
Private Sub cmdDiscount_Click()
CTR = 0

picDiscount.Print "# of Items"; "   "; " Discount"


 Open App.Path & "\Discount.txt" For Input As #1
    
    Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, ItemsPurchased(CTR), PercentageDiscount(CTR)
    picDiscount.Print ItemsPurchased(CTR), FormatPercent(PercentageDiscount(CTR))
    Loop
    Close #1
    
    

End Sub

'The shop button will allow the user to return to shopping after
'reviewing the discount information.
Private Sub cmdShop_Click()
    frmDiscount.Visible = False
    frmWomen.Visible = True
End Sub


