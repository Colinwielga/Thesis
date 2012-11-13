VERSION 5.00
Begin VB.Form frmCheckout 
   BackColor       =   &H00000000&
   Caption         =   "Checkout"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBrownBag 
      Height          =   1695
      Left            =   2880
      Picture         =   "frmCheckout.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.PictureBox picBloom4 
      Height          =   375
      Left            =   6840
      Picture         =   "frmCheckout.frx":59B2
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   6120
      Width           =   2295
   End
   Begin VB.PictureBox picBloom3 
      Height          =   375
      Left            =   4560
      Picture         =   "frmCheckout.frx":8618
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   6120
      Width           =   2295
   End
   Begin VB.PictureBox PicBloom2 
      Height          =   375
      Left            =   2280
      Picture         =   "frmCheckout.frx":B27E
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   6120
      Width           =   2295
   End
   Begin VB.PictureBox PicBloom1 
      Height          =   375
      Left            =   0
      Picture         =   "frmCheckout.frx":DEE4
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C000&
      Caption         =   "Thank You for Your  Purchase! Goodbye!"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00C000C0&
      Caption         =   "Calculate Total Purchase"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5160
      ScaleHeight     =   2835
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   1680
      Width           =   3735
   End
End
Attribute VB_Name = "frmCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form will calculate the total for your purchase,
'including sales tax, and apply any discuonts that may apply.
'Then, the total will be printed with the correct discouoned calculated.

'Declare all variables needed for commands
Dim Total As Single

'Exit theprogram
Private Sub cmdQuit_Click()
End
End Sub

'The total button will use the ShoppingCart and Items variables
'as indicators of the purchase amount and discount amount.  It will
'then add a sales tax of 5% and calculate the discount based on the
'number of items purchased.
Private Sub cmdTotal_Click()


    If Items = 2 Then
        Total = ((ShoppingCart * 1.05) - (ShoppingCart * 0.05))
        picResults.Print "Your total is:  "; FormatCurrency(Total)
    ElseIf Items = 3 Then
        Total = ((ShoppingCart * 1.05) - (ShoppingCart * 0.08))
        picResults.Print "Your total is  :"; FormatCurrency(Total)
    ElseIf Items = 4 Then
        Total = ((ShoppingCart * 1.05) - (ShoppingCart * 0.1))
        picResults.Print "Your total is:  "; FormatCurrency(Total)
    ElseIf Items = 5 Then
        Total = ((ShoppingCart * 1.05) - (ShoppingCart * 0.15))
        picResults.Print "Your total is:  "; FormatCurrency(Total)
    ElseIf Items = 6 Then
        Total = ((ShoppingCart * 1.05) - (ShoppingCart * 0.18))
        picResults.Print "Your total is:  "; FormatCurrency(Total)
    ElseIf Items < 2 Then
        Total = ((ShoppingCart * 1.05))
        picResults.Print "Your total is:  "; FormatCurrency(Total)
    ElseIf Items > 6 Then
        Total = ((ShoppingCart * 1.05))
        picResults.Print "Your total is:  "; FormatCurrency(Total)
    End If

End Sub
