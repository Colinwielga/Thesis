VERSION 5.00
Begin VB.Form frmLeotards 
   BackColor       =   &H00FF80FF&
   Caption         =   "Leotards"
   ClientHeight    =   9840
   ClientLeft      =   885
   ClientTop       =   540
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   12780
   Visible         =   0   'False
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   735
      Left            =   4080
      TabIndex        =   16
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   735
      Left            =   2760
      TabIndex        =   15
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   1440
      TabIndex        =   14
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdLeotard8 
      Caption         =   "Peasant Top Cap Sleeve  $36.00"
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton cmdLeotard7 
      Caption         =   "Mock Turtleneck Halter Leotard  $26.25"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   8400
      Width           =   2535
   End
   Begin VB.CommandButton cmdLeotard6 
      Caption         =   "Double Strap Camisole Leotard  $24.75"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton cmdLeotard4 
      Caption         =   "Fan Back Camisole  $26.25"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmdLeotard3 
      Caption         =   "Cap Sleeve Leotard  $24.45"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmdLeotard1 
      Caption         =   "3/4 Sleeve Leotard  $25.50"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   6375
      Left            =   8520
      ScaleHeight     =   6315
      ScaleWidth      =   4035
      TabIndex        =   7
      Top             =   120
      Width           =   4095
   End
   Begin VB.PictureBox Picture10 
      Height          =   4095
      Left            =   5880
      Picture         =   "Leotards.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   4200
      Width           =   2175
   End
   Begin VB.PictureBox Picture9 
      Height          =   4095
      Left            =   3240
      Picture         =   "Leotards.frx":60EF
      ScaleHeight     =   4035
      ScaleWidth      =   1755
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.PictureBox Picture8 
      Height          =   3495
      Left            =   5880
      Picture         =   "Leotards.frx":BBF4
      ScaleHeight     =   3435
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox Picture7 
      Height          =   4095
      Left            =   120
      Picture         =   "Leotards.frx":11D6E
      ScaleHeight     =   4035
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
   End
   Begin VB.PictureBox Picture5 
      Height          =   3495
      Left            =   3000
      Picture         =   "Leotards.frx":17D42
      ScaleHeight     =   3435
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   120
      Picture         =   "Leotards.frx":1E3A9
      ScaleHeight     =   3435
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   10320
      TabIndex        =   18
      Top             =   9360
      Width           =   2295
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H00C000C0&
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
      Height          =   2415
      Left            =   8280
      TabIndex        =   17
      Top             =   6840
      Width           =   4335
   End
End
Attribute VB_Name = "frmLeotards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmLeotards (Leotards.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy dance leotards
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
    frmLeotards.Hide
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmLeotards.Hide
End Sub

Private Sub cmdClear_Click()
TotalLeotards = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "******************************************************************************************************"

End Sub

Private Sub cmdLeotard1_Click()
Dim Leotard1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (25.5 * 0.7)
    Else
        Price = Quantity * 25.5
    End If
picResults.Print "3/4 Sleeve Leotard"; Tab(33); Quantity; Tab(45); FormatCurrency(Price)
TotalLeotards = TotalLeotards + Price
End Sub

Private Sub cmdLeotard2_Click()
Dim Leotard2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (27.45 * 0.7)
    Else
        Price = Quantity * 27.45
    End If
picResults.Print "Adjustable Strap Camisole"; Tab(33); Quantity; Tab(45); FormatCurrency(Price)
TotalLeotards = TotalLeotards + Price
End Sub

Private Sub cmdLeotard3_Click()
Dim Leotard3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (24.45 * 0.7)
    Else
        Price = Quantity * 24.45
    End If
picResults.Print "Cap Sleeve Leotard"; Tab(33); Quantity; Tab(45); FormatCurrency(Price)
TotalLeotards = TotalLeotards + Price
End Sub

Private Sub cmdLeotard4_Click()
Dim Leotard4 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (26.25 * 0.7)
    Else
        Price = Quantity * 26.25
    End If
picResults.Print "Fan Back Camisole"; Tab(33); Quantity; Tab(45); FormatCurrency(Price)
TotalLeotards = TotalLeotards + Price
End Sub

Private Sub cmdLeotard5_Click()
Dim Leotard5 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (29.65 * 0.7)
    Else
        Price = Quantity * 29.65
    End If
picResults.Print "Zipper Front Halter"; Tab(33); Quantity; Tab(45); FormatCurrency(Price)
TotalLeotards = TotalLeotards + Price
End Sub

Private Sub cmdLeotard6_Click()
Dim Leotard6 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (24.75 * 0.7)
    Else
        Price = Quantity * 24.75
    End If
picResults.Print "Double Strap Camisole Leotard"; Tab(33); Quantity; Tab(45); FormatCurrency(Price)
TotalLeotards = TotalLeotards + Price
End Sub

Private Sub cmdLeotard7_Click()
Dim Leotard7 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (26.25 * 0.7)
    Else
        Price = Quantity * 26.25
    End If
picResults.Print "Mock Turtleneck Halter Leotard"; Tab(33); Quantity; Tab(45); FormatCurrency(Price)
TotalLeotards = TotalLeotards + Price
End Sub

Private Sub cmdLeotard8_Click()
Dim Leotard8 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (36 * 0.7)
    Else
        Price = Quantity * 36
    End If
picResults.Print "Peasant Top Cap Sleeve"; Tab(33); Quantity; Tab(45); FormatCurrency(Price)
Total = Total + Price
End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "**************************************************************************************************************"
picResults.Print "SubTotal"; Tab(45); FormatCurrency(TotalLeotards)
tax = TotalLeotards * 0.07
picResults.Print "Tax"; Tab(45); FormatCurrency(tax)
TotalLeotards = TotalLeotards + tax
picResults.Print "Total"; Tab(45); FormatCurrency(TotalLeotards)
End Sub


