VERSION 5.00
Begin VB.Form frmAccessories 
   BackColor       =   &H00800000&
   Caption         =   "Accessories"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   150
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   14925
   Visible         =   0   'False
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   4920
      TabIndex        =   15
      Top             =   9840
      Width           =   1095
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2520
      TabIndex        =   14
      Top             =   9840
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1320
      TabIndex        =   13
      Top             =   9840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAccessories5 
      Caption         =   "Point Comfort Toe Pads  $17.20"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton cmdAccessories4 
      Caption         =   "Blister-Aid Kit  $10.50"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   9360
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccessories3 
      Caption         =   "Toe Savers Prima Toe Pad  $19.05"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   9360
      Width           =   2655
   End
   Begin VB.CommandButton cmdAccessories2 
      Caption         =   "Gel Pointe Shoe Cushions  $1.90"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton cmdAccessories1 
      Caption         =   "Capezio Toe Pad  $4.25"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next Page"
      Height          =   735
      Left            =   3720
      TabIndex        =   7
      Top             =   9840
      Width           =   1095
   End
   Begin VB.PictureBox Picture8 
      Height          =   4095
      Left            =   3600
      Picture         =   "Accessories.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   3075
      TabIndex        =   6
      Top             =   5160
      Width           =   3135
   End
   Begin VB.PictureBox Picture6 
      Height          =   3615
      Left            =   6720
      Picture         =   "Accessories.frx":524A
      ScaleHeight     =   3555
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.PictureBox Picture4 
      Height          =   3615
      Left            =   3480
      Picture         =   "Accessories.frx":CB04
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   120
      Picture         =   "Accessories.frx":10EA1
      ScaleHeight     =   3555
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   120
      Picture         =   "Accessories.frx":15CBD
      ScaleHeight     =   4875
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   4320
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   8175
      Left            =   10560
      ScaleHeight     =   8115
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   9840
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   12600
      TabIndex        =   17
      Top             =   10200
      Width           =   2175
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H00FFC0C0&
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
      Height          =   2295
      Left            =   6960
      TabIndex        =   16
      Top             =   5040
      Width           =   3495
   End
End
Attribute VB_Name = "frmAccessories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmAccessories (Accessories.frm)
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

Private Sub cmdAccessories1_Click()
Dim Accessories1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (4.25 * 0.7)
    Else
        Price = Quantity * 4.25
    End If
picResults.Print "Capezio Toe Pad"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories = TotalAccessories + Price
End Sub

Private Sub cmdAccessories2_Click()
Dim Accessories2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (1.9 * 0.7)
    Else
        Price = Quantity * 1.9
    End If
picResults.Print "Gel Pointe Shoe Cushions"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories = TotalAccessories + Price
End Sub

Private Sub cmdAccessories3_Click()
Dim Accessories3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (19.05 * 0.7)
    Else
        Price = Quantity * 19.05
    End If
picResults.Print "Toe Savers Prima Toe Pad"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories = TotalAccessories + Price
End Sub

Private Sub cmdAccessories4_Click()
Dim Accessories4 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (10.5 * 0.7)
    Else
        Price = Quantity * 10.5
    End If
picResults.Print "Blister-Aid Kit"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories = TotalAccessories + Price
End Sub

Private Sub cmdAccessories5_Click()
Dim Accessories5 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (17.2 * 0.7)
    Else
        Price = Quantity * 17.2
    End If
picResults.Print "Point Comfort Toe Pads"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalAccessories = TotalAccessories + Price
End Sub

Private Sub cmdBack_Click()
    frmShoesetc.Show
    frmAccessories.Hide
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmAccessories.Hide
End Sub

Private Sub cmdClear_Click()
TotalAccessories = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "************************************************************************************"

End Sub

Private Sub cmdNext_Click()
    frmAccessories2.Show
    frmAccessories.Hide
    frmAccessories2.picResults.Cls
    frmAccessories2.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmAccessories2.picResults.Print "******************************************************************************************************"
End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "*********************************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalAccessories)
tax = TotalAccessories * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalAccessories = TotalAccessories + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalAccessories)
End Sub

