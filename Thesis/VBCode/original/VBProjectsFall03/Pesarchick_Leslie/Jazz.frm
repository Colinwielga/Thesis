VERSION 5.00
Begin VB.Form frmJazz 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   9930
   ClientLeft      =   255
   ClientTop       =   735
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   13995
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   3720
      TabIndex        =   14
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2520
      TabIndex        =   13
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1320
      TabIndex        =   12
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdJazz5 
      Caption         =   "Split-Sole Dance Sneaker  $49.50"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   8520
      Width           =   2775
   End
   Begin VB.CommandButton cmdJazz4 
      Caption         =   "Leather Dansneakers  $52.95"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   8520
      Width           =   2415
   End
   Begin VB.CommandButton cmdJazz3 
      Caption         =   "Ultraflex Jazz Shoe  $32.25"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdJazz2 
      Caption         =   "Economy Jazz Shoe  $24.45"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdJazz1 
      Caption         =   "Classic Jazz Boot  $45.95"
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
      Top             =   9000
      Width           =   1095
   End
   Begin VB.PictureBox Picture6 
      Height          =   3855
      Left            =   6360
      Picture         =   "Jazz.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture5 
      Height          =   3855
      Left            =   3240
      Picture         =   "Jazz.frx":6002
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   4560
      Width           =   3015
   End
   Begin VB.PictureBox Picture4 
      Height          =   3855
      Left            =   120
      Picture         =   "Jazz.frx":D870
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   4560
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Height          =   3855
      Left            =   3240
      Picture         =   "Jazz.frx":14460
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   120
      Picture         =   "Jazz.frx":1A113
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
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblName 
      BackColor       =   &H008080FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   11640
      TabIndex        =   16
      Top             =   9360
      Width           =   2175
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H008080FF&
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
      Height          =   4455
      Left            =   6480
      TabIndex        =   15
      Top             =   4920
      Width           =   2775
   End
End
Attribute VB_Name = "frmJazz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmJazz (Jazz.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy Jazz Shoes
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
    frmJazz.Hide
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmJazz.Hide
End Sub

Private Sub cmdClear_Click()
TotalJazz = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "***********************************************************************************************************"
End Sub

Private Sub cmdJazz1_Click()
Dim Jazz1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (45.95 * 0.7)
    Else
        Price = Quantity * 45.95
    End If
picResults.Print "Classic Jazz Boot"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalJazz = TotalJazz + Price
End Sub

Private Sub cmdJazz2_Click()
Dim Jazz2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (24.45 * 0.7)
    Else
        Price = Quantity * 24.45
    End If
picResults.Print "Economy Jazz Shoe"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalJazz = TotalJazz + Price
End Sub

Private Sub cmdJazz3_Click()
Dim Jazz3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (32.25 * 0.7)
    Else
        Price = Quantity * 32.25
    End If
picResults.Print "Ultraflex Jazz Shoe"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalJazz = TotalJazz + Price
End Sub

Private Sub cmdJazz4_Click()
Dim Jazz4 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (52.95 * 0.7)
    Else
        Price = Quantity * 52.95
    End If
Jazz4 = 52.95
picResults.Print "Leather Dansneakers"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalJazz = TotalJazz + Price
End Sub

Private Sub cmdJazz5_Click()
Dim Jazz5 As Single
Jazz5 = 49.5
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (49.5 * 0.7)
    Else
        Price = Quantity * 49.5
    End If
picResults.Print "Split-Sole Dance Sneaker"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalJazz = TotalJazz + Price
End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "**********************************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalJazz)
tax = TotalJazz * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalJazz = TotalJazz + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalJazz)
End Sub

