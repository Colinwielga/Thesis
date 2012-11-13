VERSION 5.00
Begin VB.Form frmStore
   Caption         =   "Form1"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   10125
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear
      Caption         =   "Clear"
      Height          =   975
      Left            =   9600
      TabIndex        =   22
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdbs
      Caption         =   "Back"
      Height          =   975
      Left            =   11400
      TabIndex        =   21
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdTotal
      Caption         =   "Total"
      Height          =   975
      Left            =   7920
      TabIndex        =   20
      Top             =   8400
      Width           =   1335
   End
   Begin VB.PictureBox picResults
      Height          =   6735
      Left            =   7920
      ScaleHeight     =   6675
      ScaleWidth      =   4515
      TabIndex        =   18
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CommandButton cmd9
      Height          =   1695
      Left            =   5280
      Picture         =   "frmStore.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmd8
      Height          =   1695
      Left            =   3000
      Picture         =   "frmStore.frx":18ED
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmd7
      Height          =   1695
      Left            =   360
      Picture         =   "frmStore.frx":2BA3
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmd6
      Height          =   1935
      Left            =   5160
      Picture         =   "frmStore.frx":3ECC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton cmd5
      Height          =   1575
      Left            =   3000
      Picture         =   "frmStore.frx":579B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmd4
      Height          =   1575
      Left            =   360
      Picture         =   "frmStore.frx":6AB8
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmd3
      Height          =   1815
      Left            =   5160
      Picture         =   "frmStore.frx":797E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmd2
      Height          =   1695
      Left            =   3000
      Picture         =   "frmStore.frx":8ACF
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmd1
      Height          =   1695
      Left            =   360
      Picture         =   "frmStore.frx":9906
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblshopping
      BackStyle       =   0  'Transparent
      Caption         =   "Shopping Cart"
      BeginProperty Font
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   8520
      TabIndex        =   19
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lbl9
      BackStyle       =   0  'Transparent
      Caption         =   "Kick ass Zip Hoodie ($43.99)"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label lbl8
      BackStyle       =   0  'Transparent
      Caption         =   "Kick ass Hat    ($15.99)"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lbl7
      BackStyle       =   0  'Transparent
      Caption         =   "Kick ass Travel mug  ($7.99)"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label lbl6
      BackStyle       =   0  'Transparent
      Caption         =   "Kick ass Mug   ($8.99)"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lbl5
      BackStyle       =   0  'Transparent
      Caption         =   "Kick ass mousepad ($10.99)"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lbl4
      BackStyle       =   0  'Transparent
      Caption         =   "Kick ass light Tshirt                 ($14.99)"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lbl3
      BackStyle       =   0  'Transparent
      Caption         =   "Kick ass superhero Ringer T ($15.99)"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lbl2
      BackStyle       =   0  'Transparent
      Caption         =   "Kick ass superhero JR     ($15.99)"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lbl1
      BackStyle       =   0  'Transparent
      Caption         =   "2.25 Button ($5.59)"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image1
      BorderStyle     =   1  'Fixed Single
      Height          =   10575
      Left            =   -720
      Picture         =   "frmStore.frx":B498
      Top             =   -480
      Width           =   27060
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim total As Double
Private Sub cmd1_Click()
    ' comment every few lines
total = total + 5.59
picResults.Print "2.25 Button"; Tab(20); FormatCurrency(5.59)
End Sub
    ' comment every few lines

Private Sub cmd2_Click()
    ' comment every few lines
total = total + 15.99
picResults.Print "Kick ass JR"; Tab(20); FormatCurrency(15.99)
End Sub

Private Sub cmd3_Click()
total = total + 15.99
    ' comment every few lines
picResults.Print "Kick ass Ringer T"; Tab(20); FormatCurrency(15.99)
End Sub

Private Sub cmd4_Click()
total = total + 14.99
    ' comment every few lines
picResults.Print "Kick ass light T shirt"; Tab(20); FormatCurrency(14.99)
End Sub

Private Sub cmd5_Click()
total = total + 10.99
picResults.Print "Kick ass mousepad"; Tab(20); FormatCurrency(10.99)
End Sub

Private Sub cmd6_Click()
    ' comment every few lines
total = total + 8.99
picResults.Print "Kick ass Mug"; Tab(20); FormatCurrency(8.99)
End Sub

Private Sub cmd7_Click()
    ' comment every few lines
total = total + 7.99
picResults.Print "Kick ass travelMug"; Tab(20); FormatCurrency(7.99)
End Sub

    ' comment every few lines
Private Sub cmd8_Click()
total = total + 15.99
picResults.Print "Kick ass Hat"; Tab(20); FormatCurrency(15.99)
End Sub
    ' comment every few lines
    ' comment every few lines

Private Sub cmd9_Click()
total = total + 43.99
picResults.Print "Kick ass Zip Hoodie"; Tab(20); FormatCurrency(43.99)
End Sub

Private Sub cmdbs_Click()
frmStore.Hide
    ' comment every few lines
frmMain.Show
End Sub

Private Sub cmdclear_Click()
picResults.Cls
total = 0
End Sub

Private Sub cmdTotal_Click()
    ' comment every few lines
Dim Tax As Double
Tax = total * 0.07
    ' comment every few lines
picResults.Print "*********************************************************"
picResults.Print "Subtotal"; Tab(20); FormatCurrency(total)
picResults.Print "Tax"; Tab(20); FormatCurrency(Tax)
picResults.Print "Total"; Tab(20); FormatCurrency(Tax + total)
    ' comment every few lines
 If total > 200 Then
            MsgBox "Hey, because your purchase over 200 dollars, I will give you a small gift."
        End If
    ' comment every few lines
End Sub

