VERSION 5.00
Begin VB.Form frmUnitards 
   BackColor       =   &H00FF8080&
   Caption         =   "Unitards"
   ClientHeight    =   9150
   ClientLeft      =   -135
   ClientTop       =   -45
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   15240
   Visible         =   0   'False
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   4080
      TabIndex        =   12
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2760
      TabIndex        =   11
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1440
      TabIndex        =   10
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnitard4 
      Caption         =   "Sleeveless Jumper with Zipper Front  $53.25"
      Height          =   615
      Left            =   8280
      TabIndex        =   9
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton cmdUnitard3 
      Caption         =   "Hologram Mock Turtlenck Unitard  $49.90"
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CommandButton cmdUnitard2 
      Caption         =   "Mock Turtleneck Sleeveless Jazz Unitard  $29.95"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   6960
      Width           =   3375
   End
   Begin VB.CommandButton cmdUnitard1 
      Caption         =   "Ballerina Bodice Unitard  $32.65"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   7455
      Left            =   11160
      ScaleHeight     =   7395
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   120
      Width           =   4695
   End
   Begin VB.PictureBox Picture5 
      Height          =   6735
      Left            =   8280
      Picture         =   "Unitards.frx":0000
      ScaleHeight     =   6675
      ScaleWidth      =   2715
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox Picture4 
      Height          =   6735
      Left            =   2280
      Picture         =   "Unitards.frx":2A53
      ScaleHeight     =   6675
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin VB.PictureBox Picture3 
      Height          =   6735
      Left            =   5760
      Picture         =   "Unitards.frx":5CEC
      ScaleHeight     =   6675
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   6735
      Left            =   120
      Picture         =   "Unitards.frx":8C29
      ScaleHeight     =   6675
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   13560
      TabIndex        =   14
      Top             =   8400
      Width           =   2295
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
      Height          =   1215
      Left            =   5400
      TabIndex        =   13
      Top             =   7680
      Width           =   7335
   End
End
Attribute VB_Name = "frmUnitards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmUnitards (Unitards.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy unitards
                    'if they buy over 20 items, they receive 30% off
                    'totals what they buy, and adds a 7% tax
                    'prints out total on this form, and on frmshoesetc

Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.
Dim Quantity As Integer
Dim Price As Single
Private Sub cmdBack_Click()
    frmShoesetc.Show
    frmUnitards.Hide
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmUnitards.Hide
End Sub

Private Sub cmdClear_Click()
TotalUnitards = 0
picResults.Cls
picResults.Print "Item"; Tab(43); "Quantity"; Tab(50); "Price"
picResults.Print "************************************************************************************"

End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "***********************************************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalUnitards)
tax = TotalUnitards * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalUnitards = TotalUnitards + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalUnitards)
End Sub

Private Sub cmdUnitard1_Click()
Dim Unitard1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (32.65 * 0.7)
    Else
        Price = Quantity * 32.65
    End If
picResults.Print "Ballerina Bodice Unitard"; Tab(43); Quantity; Tab(50); FormatCurrency(Price)
TotalUnitards = TotalUnitards + Price
End Sub

Private Sub cmdUnitard2_Click()
Dim Unitard2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (29.95 * 0.7)
    Else
        Price = Quantity * 29.95
    End If
picResults.Print "Mock Turtleneck Sleeveless Jazz Unitard"; Tab(43); Quantity; Tab(50); FormatCurrency(Price)
TotalUnitards = TotalUnitards + Price
End Sub

Private Sub cmdUnitard3_Click()
Dim Unitard3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (49.9 * 0.7)
    Else
        Price = Quantity * 49.9
    End If
picResults.Print "Hologram Mock Turtleneck Unitard"; Tab(43); Quantity; Tab(50); FormatCurrency(Price)
TotalUnitards = TotalUnitards + Price
End Sub

Private Sub cmdUnitard4_Click()
Dim Unitard4 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (53.25 * 0.7)
    Else
        Price = Quantity * 53.25
    End If
picResults.Print "Sleeveless Jumper with Zipper Front"; Tab(43); Quantity; Tab(50); FormatCurrency(Price)
TotalUnitards = TotalUnitards + Price
End Sub
