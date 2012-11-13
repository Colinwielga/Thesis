VERSION 5.00
Begin VB.Form frmBallet 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   13905
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   4920
      TabIndex        =   15
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2520
      TabIndex        =   14
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1320
      TabIndex        =   13
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdPoint5 
      Caption         =   "Plie II  $45.90"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton cmdPoint4 
      Caption         =   "Alpha Pointe Shoe  $48.00"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton cmdPoint3 
      Caption         =   "Ulanova  $47.90"
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdPoint2 
      Caption         =   "Glisse  $42.90"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdPoint1 
      Caption         =   "Lyrica  $28.10"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Page"
      Height          =   735
      Left            =   3720
      TabIndex        =   7
      Top             =   9000
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   7095
      Left            =   9600
      ScaleHeight     =   7035
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   120
      Width           =   4095
   End
   Begin VB.PictureBox Picture5 
      Height          =   3855
      Left            =   6360
      Picture         =   "Ballet.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture4 
      Height          =   3855
      Left            =   120
      Picture         =   "Ballet.frx":5401
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Height          =   3855
      Left            =   3240
      Picture         =   "Ballet.frx":9B58
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   3240
      Picture         =   "Ballet.frx":EA4F
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   4560
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   120
      Picture         =   "Ballet.frx":13052
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   4560
      Width           =   3015
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   11400
      TabIndex        =   17
      Top             =   9240
      Width           =   2295
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H00FF80FF&
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
      Height          =   3495
      Left            =   6480
      TabIndex        =   16
      Top             =   4920
      Width           =   2895
   End
End
Attribute VB_Name = "frmBallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmBallet (Ballet.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy ballet shoes
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
    frmBallet.Hide
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmBallet.Hide
End Sub

Private Sub cmdClear_Click()
TotalBallet = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "******************************************************************************************************"

End Sub

Private Sub cmdNext_Click()
    frmBallet2.Show
    frmBallet.Hide
    frmBallet2.picResults.Cls
    frmBallet2.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmBallet2.picResults.Print "************************************************************************************"
End Sub

Private Sub cmdPoint1_Click()
Dim Point1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (28.1 * 0.7)
    Else
        Price = Quantity * 28.1
    End If
picResults.Print "Lyrica"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet = TotalBallet + Price
End Sub

Private Sub cmdPoint2_Click()
Dim Point2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (42.9 * 0.7)
    Else
        Price = Quantity * 42.9
    End If
picResults.Print "Glisse"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet = TotalBallet + Price
End Sub

Private Sub cmdPoint3_Click()
Dim Point3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (47.9 * 0.7)
    Else
        Price = Quantity * 47.9
    End If
picResults.Print "Ulanova"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet = TotalBallet + Price
End Sub

Private Sub cmdPoint4_Click()
Dim Point4 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (48 * 0.7)
    Else
        Price = Quantity * 48
    End If
picResults.Print "Alpha Pointe Shoe"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet = TotalBallet + Price
End Sub

Private Sub cmdPoint5_Click()
Dim Point5 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (45.9 * 0.7)
    Else
        Price = Quantity * 45.9
    End If
picResults.Print "Plie II"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalBallet = TotalBallet + Price
End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "*********************************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalBallet)
tax = TotalBallet * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalBallet = TotalBallet + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalBallet)
End Sub

