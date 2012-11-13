VERSION 5.00
Begin VB.Form frmBallet2 
   BackColor       =   &H00404080&
   Caption         =   "Form1"
   ClientHeight    =   9855
   ClientLeft      =   255
   ClientTop       =   540
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   13890
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   4200
      TabIndex        =   14
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2880
      TabIndex        =   13
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1560
      TabIndex        =   12
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdBallet5 
      Caption         =   "Prolite II Ballet Slipper  $21.55"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   8520
      Width           =   2415
   End
   Begin VB.CommandButton cmdBallet4 
      Caption         =   "Leather Split-Sole Ballet Slipper  $25.50"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   8520
      Width           =   3015
   End
   Begin VB.CommandButton cmdBallet3 
      Caption         =   "Dansoft Ballet Slipper  $21.00"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdBallet2 
      Caption         =   "Ultimate  $13.20"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdBallet1 
      Caption         =   "Daisy Ballet Slipper  $12.00"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   6735
      Left            =   9600
      ScaleHeight     =   6675
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   120
      Width           =   4095
   End
   Begin VB.PictureBox Picture6 
      Height          =   3855
      Left            =   3240
      Picture         =   "Ballet2.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture5 
      Height          =   3855
      Left            =   3240
      Picture         =   "Ballet2.frx":6CF5
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   4560
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Height          =   3855
      Left            =   120
      Picture         =   "Ballet2.frx":B8A3
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   4560
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   6360
      Picture         =   "Ballet2.frx":10DE1
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   120
      Picture         =   "Ballet2.frx":15936
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Previous Page"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   11520
      TabIndex        =   16
      Top             =   9240
      Width           =   2175
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H00C0C0FF&
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
      Height          =   4215
      Left            =   6480
      TabIndex        =   15
      Top             =   4920
      Width           =   2895
   End
End
Attribute VB_Name = "frmBallet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmBallet2 (Ballet2.frm)
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
Private Sub cmdBack_Click()
    frmBallet.Show
    frmBallet2.Hide
    frmBallet.picResults.Cls
    frmBallet.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmBallet.picResults.Print "*********************************************************************************"
End Sub
Private Sub cmdBallet1_Click()
Dim Ballet1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (12 * 0.7)
    Else
        Price = Quantity * 12
    End If
picResults.Print "Daisy Ballet Slipper"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet2 = TotalBallet2 + Price
End Sub

Private Sub cmdBallet2_Click()
Dim Ballet2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (13.2 * 0.7)
    Else
        Price = Quantity * 13.2
    End If
picResults.Print "Ultimate"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet2 = TotalBallet2 + Price
End Sub

Private Sub cmdBallet3_Click()
Dim Ballet3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (21 * 0.7)
    Else
        Price = Quantity * 21
    End If
picResults.Print "Dansoft Ballet Slipper"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet2 = TotalBallet2 + Price
End Sub

Private Sub cmdBallet4_Click()
Dim Ballet4 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (25.5 * 0.7)
    Else
        Price = Quantity * 25.5
    End If
picResults.Print "Leather Split-Sole Ballet Slipper"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet2 = TotalBallet2 + Price
End Sub

Private Sub cmdBallet5_Click()
Dim Ballet5 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (21.55 * 0.7)
    Else
        Price = Quantity * 21.55
    End If
picResults.Print "Prolite II Ballet Slipper"; Tab(35); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet2 = TotalBallet2 + Price
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmBallet2.Hide
End Sub

Private Sub cmdClear_Click()
TotalBallet2 = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "************************************************************************************"
End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "*****************************************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalBallet2)
tax = TotalBallet2 * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalBallet2 = TotalBallet2 + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalBallet2)
End Sub
