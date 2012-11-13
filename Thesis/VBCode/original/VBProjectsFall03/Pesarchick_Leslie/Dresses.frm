VERSION 5.00
Begin VB.Form frmDresses 
   BackColor       =   &H00FFFF00&
   Caption         =   "Dresses"
   ClientHeight    =   11100
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   ScaleHeight     =   11100
   ScaleWidth      =   13365
   Visible         =   0   'False
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   3840
      TabIndex        =   16
      Top             =   10200
      Width           =   1095
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2640
      TabIndex        =   15
      Top             =   10200
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1440
      TabIndex        =   14
      Top             =   10200
      Width           =   1095
   End
   Begin VB.CommandButton cmdDress6 
      Caption         =   "Sheer Overdress  $31.50"
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton cmdDress5 
      Caption         =   "Empire Waist Dance Dress  $41.25"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   9720
      Width           =   2775
   End
   Begin VB.CommandButton cmdDress4 
      Caption         =   "Velvet Strap Camisole Dress  $40.15"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   9720
      Width           =   2775
   End
   Begin VB.CommandButton cmdDress3 
      Caption         =   "Velvet Fringe Dress  $45.00"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdDress2 
      Caption         =   "Drop Waist Dress w/wings  $94.50"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton cmdDress1 
      Caption         =   "Camisole Empire Dress  $78.00"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   7335
      Left            =   9120
      ScaleHeight     =   7275
      ScaleWidth      =   3795
      TabIndex        =   7
      Top             =   120
      Width           =   3855
   End
   Begin VB.PictureBox Picture7 
      Height          =   4455
      Left            =   6480
      Picture         =   "Dresses.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox Picture6 
      Height          =   3615
      Left            =   3240
      Picture         =   "Dresses.frx":5EA6
      ScaleHeight     =   3555
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   5520
      Width           =   3015
   End
   Begin VB.PictureBox Picture4 
      Height          =   4455
      Left            =   6360
      Picture         =   "Dresses.frx":B5AD
      ScaleHeight     =   4395
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.PictureBox Picture3 
      Height          =   4215
      Left            =   240
      Picture         =   "Dresses.frx":11A78
      ScaleHeight     =   4155
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   5280
      Width           =   2895
   End
   Begin VB.PictureBox Picture2 
      Height          =   4455
      Left            =   3360
      Picture         =   "Dresses.frx":17EB2
      ScaleHeight     =   4395
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   240
      Picture         =   "Dresses.frx":1FB27
      ScaleHeight     =   4395
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   10200
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   10800
      TabIndex        =   18
      Top             =   10560
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
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
      Height          =   2535
      Left            =   8880
      TabIndex        =   17
      Top             =   7920
      Width           =   4335
   End
End
Attribute VB_Name = "frmDresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmDresses (Dresses.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy dresses
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
    frmDresses.Hide
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmDresses.Hide
End Sub

Private Sub cmdClear_Click()
TotalDresses = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "************************************************************************************"

End Sub

Private Sub cmdDress1_Click()
Dim Dress1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (78 * 0.7)
    Else
        Price = Quantity * 78
    End If
picResults.Print "Camisole Empire Dress"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalDresses = TotalDresses + Price
End Sub

Private Sub cmdDress2_Click()
Dim Dress2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (94.5 * 0.7)
    Else
        Price = Quantity * 94.5
    End If
picResults.Print "Drop Waist Dress w/Wings"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalDresses = TotalDresses + Price
End Sub

Private Sub cmdDress3_Click()
Dim Dress3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (45 * 0.7)
    Else
        Price = Quantity * 45
    End If
picResults.Print "Velvet Fringe Dress"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalDresses = TotalDresses + Price
End Sub

Private Sub cmdDress4_Click()
Dim Dress4 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (40.15 * 0.7)
    Else
        Price = Quantity * 40.15
    End If
picResults.Print "Velvet Strap Camisole Dress"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalDresses = TotalDresses + Price
End Sub

Private Sub cmdDress5_Click()
Dim Dress5 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (41.25 * 0.7)
    Else
        Price = Quantity * 41.25
    End If
picResults.Print "Empire Waist Dance Dress"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalDresses = TotalDresses + Price
End Sub

Private Sub cmdDress6_Click()
Dim Dress6 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (31.5 * 0.7)
    Else
        Price = Quantity * 31.5
    End If
picResults.Print "Sheer Overdress"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalDresses = TotalDresses + Price
End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "*******************************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalDresses)
tax = TotalDresses * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalDresses = TotalDresses + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalDresses)
End Sub

