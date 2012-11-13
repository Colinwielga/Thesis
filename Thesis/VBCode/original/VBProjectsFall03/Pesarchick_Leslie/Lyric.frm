VERSION 5.00
Begin VB.Form frmLyric 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   14340
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   855
      Left            =   4200
      TabIndex        =   10
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   855
      Left            =   2880
      TabIndex        =   9
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   855
      Left            =   1560
      TabIndex        =   8
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdLyric3 
      Caption         =   "Half Sole Modern Sandal  $14.75"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdLyric2 
      Caption         =   "Hermes Sandal  $26.95"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdLyric1 
      Caption         =   "Foot Thong  $12.75"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Default         =   -1  'True
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   6840
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   3855
      Left            =   3240
      Picture         =   "Lyric.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   6360
      Picture         =   "Lyric.frx":7801
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   120
      Picture         =   "Lyric.frx":DEC8
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      Height          =   7815
      Left            =   9600
      ScaleHeight     =   7755
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FF8080&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   11760
      TabIndex        =   12
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H00FF8080&
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
      Height          =   1455
      Left            =   480
      TabIndex        =   11
      Top             =   4680
      Width           =   7335
   End
End
Attribute VB_Name = "frmLyric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmLyric (Lyric.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy Lyric, or Modern Dance shoes
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
    frmLyric.Hide
End Sub

Private Sub cmdBuy_Click()
    frmShoesetc.Show
    frmLyric.Hide
End Sub

Private Sub cmdClear_Click()
TotalLyric = 0
picResults.Cls
picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
picResults.Print "*******************************************************************************************************"

End Sub

Private Sub cmdLyric1_Click()
Dim Lyric1 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (12.75 * 0.7)
    Else
        Price = Quantity * 12.75
    End If
picResults.Print "Foot Thong"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalLyric = TotalLyric + Price
End Sub

Private Sub cmdLyric2_Click()
Dim Lyric2 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (26.95 * 0.7)
    Else
        Price = Quantity * 26.95
    End If
picResults.Print "Hermes Sandal"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalLyric = TotalLyric + Price
End Sub

Private Sub cmdLyric3_Click()
Dim Lyric3 As Single
Quantity = InputBox("Enter the Number to Purchase", "Quantity")
If Quantity >= 20 Then
        Price = Quantity * (14.75 * 0.7)
    Else
        Price = Quantity * 14.75
    End If
picResults.Print "Half Sole Modern Sandal"; Tab(30); Quantity; Tab(41); FormatCurrency(Price)
TotalLyric = TotalLyric + Price
End Sub

Private Sub cmdTotal_Click()
Dim tax As Single
picResults.Print "************************************************************************************************"
picResults.Print "SubTotal"; Tab(41); FormatCurrency(TotalLyric)
tax = TotalLyric * 0.07
picResults.Print "Tax"; Tab(41); FormatCurrency(tax)
TotalLyric = TotalLyric + tax
picResults.Print "Total"; Tab(41); FormatCurrency(TotalLyric)
End Sub

