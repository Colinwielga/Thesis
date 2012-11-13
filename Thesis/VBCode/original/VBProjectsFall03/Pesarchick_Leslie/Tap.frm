VERSION 5.00
Begin VB.Form frmTap 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   -135
   ClientTop       =   150
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   14235
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   3720
      TabIndex        =   14
      Top             =   9120
      Width           =   1095
   End
   Begin VB.CommandButton cmdTap5 
      Caption         =   "Ladies Show Tapper  $41.25"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2520
      TabIndex        =   12
      Top             =   9120
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1320
      TabIndex        =   11
      Top             =   9120
      Width           =   1095
   End
   Begin VB.CommandButton cmdTap4 
      Caption         =   "Split-Sole Tap Shoe  $55.55"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   8520
      Width           =   2535
   End
   Begin VB.CommandButton cmdTap3 
      Caption         =   "Premiere Tap Oxford  $54.25"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   8520
      Width           =   2535
   End
   Begin VB.CommandButton cmdTap2 
      Caption         =   "Giordano Jazz Tap  $45.75"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdTap1 
      Caption         =   "Economy Tap Shoe  $21.00"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   9120
      Width           =   1095
   End
   Begin VB.PictureBox Picture6 
      Height          =   3855
      Left            =   3240
      Picture         =   "Tap.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   4560
      Width           =   3015
   End
   Begin VB.PictureBox Picture5 
      Height          =   3855
      Left            =   120
      Picture         =   "Tap.frx":600C
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   4560
      Width           =   3015
   End
   Begin VB.PictureBox Picture4 
      Height          =   3855
      Left            =   6360
      Picture         =   "Tap.frx":BD92
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Height          =   3855
      Left            =   3240
      Picture         =   "Tap.frx":11B09
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   120
      Picture         =   "Tap.frx":18941
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   7695
      Left            =   9480
      ScaleHeight     =   7635
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   11760
      TabIndex        =   16
      Top             =   9480
      Width           =   2295
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H00FFFF80&
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
      Top             =   5160
      Width           =   2655
   End
End
Attribute VB_Name = "frmTap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmTap (Tap.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy Tap Shoes
                    'if they buy over 20 pairs, they receive 30% off
                    'totals what they buy, and adds a 7% tax
                    'prints out total on this form, and on frmshoesetc

Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.
Dim Quantity As Integer
Dim Price As Single
Private Sub cmdBack_Click()
    frmShoes.Show
    frmTap.Hide
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmTap.Hide
End Sub

Private Sub cmdClear_Click()
TotalTap = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "***********************************************************************************************************"

End Sub

Private Sub cmdTap1_Click()
Dim Tap1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (21 * 0.7)
    Else
        Price = Quantity * 21
    End If
picResults.Print "Economy Tap Shoe"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalTap = TotalTap + Price
End Sub

Private Sub cmdTap2_Click()
Dim Tap2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (45.75 * 0.7)
    Else
        Price = Quantity * 45.75
    End If
picResults.Print "Giordano Jazz Tap"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalTap = TotalTap + Price
End Sub

Private Sub cmdTap3_Click()
Dim Tap3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (54.25 * 0.7)
    Else
        Price = Quantity * 54.25
    End If
picResults.Print "Premiere Tap Oxford"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalTap = TotalTap + Price
End Sub

Private Sub cmdTap4_Click()
Dim Tap4 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (55.55 * 0.7)
    Else
        Price = Quantity * 55.55
    End If
picResults.Print "Split-Sole Tap Shoe"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalTap = TotalTap + Price
End Sub

Private Sub cmdTap5_Click()
Dim Tap5 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (41.25 * 0.7)
    Else
        Price = Quantity * 41.25
    End If
picResults.Print "Ladies Show Tapper"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalTap = TotalTap + Price
End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "***********************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalTap)
tax = TotalTap * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalTap = TotalTap + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalTap)
End Sub

