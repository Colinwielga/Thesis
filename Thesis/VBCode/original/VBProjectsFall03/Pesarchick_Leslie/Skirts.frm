VERSION 5.00
Begin VB.Form frmSkirts 
   BackColor       =   &H008080FF&
   Caption         =   "Skirts"
   ClientHeight    =   8595
   ClientLeft      =   -135
   ClientTop       =   540
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   15240
   Visible         =   0   'False
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   3720
      TabIndex        =   12
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2520
      TabIndex        =   11
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1320
      TabIndex        =   10
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSkirt4 
      Caption         =   "Short Wrap Skirt  $18.00"
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton cmdSkirt3 
      Caption         =   "Crystals Wrap Skirt  $19.90"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdSkirt2 
      Caption         =   "Chiffon Wrap Skirt  $17.60"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdSkirt1 
      Caption         =   "Chiffon Wrap Skirt  $16.95"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   7815
      Left            =   11040
      ScaleHeight     =   7755
      ScaleWidth      =   4155
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.PictureBox Picture4 
      Height          =   4575
      Left            =   7920
      Picture         =   "Skirts.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   840
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Height          =   4575
      Left            =   2280
      Picture         =   "Skirts.frx":2128
      ScaleHeight     =   4515
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   5295
      Left            =   5400
      Picture         =   "Skirts.frx":4816
      ScaleHeight     =   5235
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   5295
      Left            =   120
      Picture         =   "Skirts.frx":ACE4
      ScaleHeight     =   5235
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
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   12720
      TabIndex        =   14
      Top             =   8040
      Width           =   2535
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
      Height          =   1815
      Left            =   5040
      TabIndex        =   13
      Top             =   6000
      Width           =   5775
   End
End
Attribute VB_Name = "frmSkirts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmSkirts (Skirts.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy dance skirts
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
    frmSkirts.Hide
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmSkirts.Hide
End Sub

Private Sub cmdClear_Click()
TotalSkirts = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "************************************************************************************"

End Sub

Private Sub cmdSkirt1_Click()
Dim Skirt1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (16.95 * 0.7)
    Else
        Price = Quantity * 16.95
    End If
picResults.Print "Chiffon Wrap Skirt"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalSkirts = TotalSkirts + Price
End Sub

Private Sub cmdSkirt2_Click()
Dim Skirt2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (17.6 * 0.7)
    Else
        Price = Quantity * 17.6
    End If
picResults.Print "Chiffon Wrap Skirt"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalSkirts = TotalSkirts + Price
End Sub

Private Sub cmdSkirt3_Click()
Dim Skirt3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (19.9 * 0.7)
    Else
        Price = Quantity * 19.9
    End If
picResults.Print "Crystals Wrap Skirt"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalSkirts = TotalSkirts + Price
End Sub

Private Sub cmdSkirt4_Click()
Dim Skirt4 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (18 * 0.7)
    Else
        Price = Quantity * 18
    End If
picResults.Print "Short Wrap Skirt"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalSkirts = TotalSkirts + Price
End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "**************************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalSkirts)
tax = TotalSkirts * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalSkirts = TotalSkirts + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalSkirts)
End Sub

