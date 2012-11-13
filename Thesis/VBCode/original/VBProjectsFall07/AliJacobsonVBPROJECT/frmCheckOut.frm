VERSION 5.00
Begin VB.Form frmCheckOut 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "Click Here to Exit Program Without Placing Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   5415
   End
   Begin VB.CommandButton cmdShippingInfo 
      BackColor       =   &H00C0C000&
      Caption         =   "Click Here if you would like toVerify Order and Enter Shipping Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   5415
   End
   Begin VB.PictureBox picResults 
      Height          =   5295
      Left            =   6240
      ScaleHeight     =   5235
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00FFFF00&
      Caption         =   "Next, click Here to Display Subtotal and Calculate Final Total with Optional Shipping Costs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   5415
   End
   Begin VB.CommandButton cmdDiscount 
      BackColor       =   &H00FFFFC0&
      Caption         =   "First, Click to See if you Qualify for any Discounts!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Subtotal As Single

Private Sub cmdDiscount_Click()
Dim Discount1 As Single, Discount2 As Single

Discount1 = 0.2
Discount2 = 0.1
Select Case RunningTotal
'if the runningtotal is greater than 100 then it calculates the discount and new subtotal based on the discount amount
    Case Is > 100
        Subtotal = RunningTotal * (1 - Discount1)
        MsgBox "HORRAY! You spent enough to qualify for a discount! Your new subtotal is: " & FormatCurrency(Subtotal)
    Case Is > 75
        Subtotal = RunningTotal * (1 - Discount2)
        MsgBox "HORRAY! You spent enough to qualify for a discount! Your new subtotal is: " & FormatCurrency(Subtotal)
    Case Is > 0
        MsgBox "I'm sorry, this purchase does not qualify for a discount"
        Subtotal = RunningTotal
    End Select
    
End Sub

Private Sub cmdDisplay_Click()
Dim Shipping As Single, ShippingCost

Shipping = 0.0675
Total = Subtotal + (Subtotal * Shipping)
ShippingCost = Total - Subtotal


picResults.Print "**************************************"
picResults.Print "Subtotal is: "; FormatCurrency(Subtotal)
picResults.Print "**************************************"
picResults.Print "Shipping cost is: "; FormatCurrency(ShippingCost)
picResults.Print "**************************************"
picResults.Print "Total Cost is: "; FormatCurrency(Total)
picResults.Print "**************************************"

End Sub

Private Sub cmdShippingInfo_Click()
frmCheckOut.Hide
frmProduce.Hide
frmBakery.Hide
frmFrozen.Hide
frmEnter.Hide
frmShippingInfo.Show
End Sub

Private Sub Command1_Click()
MsgBox "Remember your subtotal is: " & FormatCurrency(Subtotal) & " and good luck with your in-store shopping!"

End

End Sub
