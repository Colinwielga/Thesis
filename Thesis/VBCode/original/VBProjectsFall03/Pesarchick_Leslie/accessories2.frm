VERSION 5.00
Begin VB.Form frmAccessories2 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   -135
   ClientTop       =   150
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   15240
   Visible         =   0   'False
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   3840
      TabIndex        =   14
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAccessories25 
      Caption         =   "Knee and Elbow Pads  $8.90"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   8760
      Width           =   3255
   End
   Begin VB.PictureBox Picture3 
      Height          =   3975
      Left            =   3960
      Picture         =   "accessories2.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   3675
      TabIndex        =   12
      Top             =   4680
      Width           =   3735
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2640
      TabIndex        =   11
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1440
      TabIndex        =   10
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAccessories24 
      Caption         =   "Teletone Taps and Super Taps  $11.35"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   8760
      Width           =   3255
   End
   Begin VB.CommandButton cmdAccessories23 
      Caption         =   "Heel Gripper  $2.30"
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdAccessories22 
      Caption         =   "Tube Sock Trio  $14.25"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   4200
      Width           =   3375
   End
   Begin VB.CommandButton cmdAccessories21 
      Caption         =   "Dance Socks  $3.85"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4200
      Width           =   2775
   End
   Begin VB.PictureBox Picture5 
      Height          =   3975
      Left            =   3240
      Picture         =   "accessories2.frx":7B57
      ScaleHeight     =   3915
      ScaleWidth      =   3795
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
   Begin VB.PictureBox Picture4 
      Height          =   3975
      Left            =   120
      Picture         =   "accessories2.frx":E9FA
      ScaleHeight     =   3915
      ScaleWidth      =   3435
      TabIndex        =   4
      Top             =   4680
      Width           =   3495
   End
   Begin VB.PictureBox Picture2 
      Height          =   3975
      Left            =   7200
      Picture         =   "accessories2.frx":167EE
      ScaleHeight     =   3915
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   120
      Picture         =   "accessories2.frx":1D6FB
      ScaleHeight     =   3915
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   8055
      Left            =   11040
      ScaleHeight     =   7995
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Previous Page"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackColor       =   &H0080FFFF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   12840
      TabIndex        =   16
      Top             =   9600
      Width           =   2175
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H0000FFFF&
      Caption         =   "30% Discount if you buy more than 20 of any item."
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   7920
      TabIndex        =   15
      Top             =   4800
      Width           =   2895
   End
End
Attribute VB_Name = "frmAccessories2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmAccessories2 (Accessories2.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy dance accessories
                    'if they buy over 20 items, they receive 30% off
                    'totals what they buy, and adds a 7% tax
                    'prints out total on this form, and on frmshoesetc
Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.
Dim Quantity As Integer
Dim Price As Single
Private Sub cmdAccessories21_Click()
Dim Accessories21 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (3.85 * 0.7)
    Else
        Price = Quantity * 3.85
    End If
picResults.Print "Dance Socks"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories2 = TotalAccessories2 + Price
End Sub


Private Sub cmdAccessories22_Click()
Dim Accessories22 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (14.25 * 0.7)
    Else
        Price = Quantity * 14.25
    End If
picResults.Print "Tube Sock Trio"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories2 = TotalAccessories2 + Price
End Sub

Private Sub cmdAccessories23_Click()
Dim Accessories23 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (2.3 * 0.7)
    Else
        Price = Quantity * 2.3
    End If
picResults.Print "Heel Gripper"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories2 = TotalAccessories2 + Price
End Sub

Private Sub cmdAccessories24_Click()
Dim Accessories24 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (11.35 * 0.7)
    Else
        Price = Quantity * 11.35
    End If
picResults.Print "Teletone Taps and Super Taps"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories2 = TotalAccessories2 + Price
End Sub

Private Sub cmdAccessories25_Click()
Dim Accessories25 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (8.9 * 0.7)
    Else
        Price = Quantity * 8.9
    End If
picResults.Print "Knee and Elbow Pads"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories2 = TotalAccessories2 + Price
End Sub

Private Sub cmdBack_Click()
    frmAccessories.Show
    frmAccessories2.Hide
    frmAccessories.picResults.Cls
    frmAccessories.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmAccessories.picResults.Print "*********************************************************************************"
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmAccessories2.Hide
End Sub

Private Sub cmdClear_Click()
TotalAccessories2 = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "************************************************************************************"

End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "********************************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalAccessories2)
tax = TotalAccessories2 * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalAccessories2 = TotalAccessories2 + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalAccessories2)
End Sub
