VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   ScaleHeight     =   12165
   ScaleWidth      =   14805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit to Store"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11280
      TabIndex        =   22
      Top             =   10800
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   5175
      Left            =   11160
      ScaleHeight     =   5115
      ScaleWidth      =   3315
      TabIndex        =   21
      Top             =   4920
      Width           =   3375
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Click to Compute Total"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11400
      TabIndex        =   20
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Information"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11280
      TabIndex        =   19
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtj6 
      Height          =   495
      Left            =   9480
      TabIndex        =   11
      Text            =   "0"
      Top             =   9120
      Width           =   975
   End
   Begin VB.TextBox txtj5 
      Height          =   495
      Left            =   9360
      TabIndex        =   10
      Text            =   "0"
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtj4 
      Height          =   495
      Left            =   9360
      TabIndex        =   9
      Text            =   "0"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtj3 
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Text            =   "0"
      Top             =   8880
      Width           =   975
   End
   Begin VB.TextBox txtj2 
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Text            =   "0"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtj1 
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Text            =   "0"
      Top             =   1920
      Width           =   975
   End
   Begin VB.PictureBox Picture6 
      Height          =   2775
      Left            =   5880
      Picture         =   "Accessories.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   8040
      Width           =   2775
   End
   Begin VB.PictureBox Picture5 
      Height          =   3255
      Left            =   5520
      Picture         =   "Accessories.frx":3906
      ScaleHeight     =   3195
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   4320
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      Height          =   2760
      Left            =   6000
      Picture         =   "Accessories.frx":56C7
      ScaleHeight     =   2700
      ScaleWidth      =   2700
      TabIndex        =   3
      Top             =   840
      Width           =   2760
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   600
      Picture         =   "Accessories.frx":8FC4
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   7680
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   240
      Picture         =   "Accessories.frx":A1CB
      ScaleHeight     =   2115
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   600
      Picture         =   "Accessories.frx":B9F6
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   10680
      X2              =   10680
      Y1              =   720
      Y2              =   11040
   End
   Begin VB.Line Line3 
      X1              =   6240
      X2              =   10680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   4680
      X2              =   4680
      Y1              =   720
      Y2              =   11040
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblPrice5 
      BackColor       =   &H0000FF00&
      Caption         =   "$149.99"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   31
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label lblPrice6 
      BackColor       =   &H0000FF00&
      Caption         =   "$23.99"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   30
      Top             =   11040
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   15
      Left            =   9120
      TabIndex        =   29
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label lblPrice4 
      BackColor       =   &H0000FF00&
      Caption         =   "$28.99"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   28
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblPrice3 
      BackColor       =   &H0000FF00&
      Caption         =   "$12.99"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   27
      Top             =   10200
      Width           =   1095
   End
   Begin VB.Label lblPrice2 
      BackColor       =   &H0000FF00&
      Caption         =   "$29.99"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label lblPrice1 
      BackColor       =   &H0000FF00&
      Caption         =   "$80.00"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   25
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblQuantity2 
      BackColor       =   &H0000FF00&
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   24
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblQuantity 
      BackColor       =   &H0000FF00&
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   23
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblGloves 
      BackColor       =   &H0000FF00&
      Caption         =   "Real Madrid Gloves"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   18
      Top             =   11040
      Width           =   2055
   End
   Begin VB.Label lblCleats 
      BackColor       =   &H0000FF00&
      Caption         =   "Real Madrid Cleats"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label lblBall 
      BackColor       =   &H0000FF00&
      Caption         =   "Real Madrid Ball"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblSocks 
      BackColor       =   &H0000FF00&
      Caption         =   "Real Madrid Socks"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   10200
      Width           =   2055
   End
   Begin VB.Label lblShorts 
      BackColor       =   &H0000FF00&
      Caption         =   "Real Madrid Shorts"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label lblShirt 
      BackColor       =   &H0000FF00&
      Caption         =   "Real Madrid Jersey"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H0000FF00&
      Caption         =   "Accessories to buy Online"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   12
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form allows the user to purchase Real Madrid genuine products.
'a picture box will show the total of the selected items.
'user can enter the quantity of items desired to be bought
Option Explicit

Dim quantityone As Integer
Dim price(1 To 6) As Single
Dim quantities(1 To 6) As Integer
Dim CTR As Integer
Dim total As Single
Dim tax As Single
Dim i As Integer, RunningTotal As Single
Dim j1 As Integer, total1 As Single
Dim j2 As Integer, total2 As Single
Dim j3 As Integer, total3 As Single
Dim j4 As Integer, total4 As Single
Dim j5 As Integer, total5 As Single
Dim j6 As Integer, total6 As Single


Private Sub cmdBack_Click()
Information.Show
Statistics.Hide
PlayersStat.Hide
Form1.Hide
OpenPage.Hide
End Sub

Private Sub cmdCompute_Click()
    PicResults.Cls
    Open App.path & "\Accessories.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, price(CTR)
    Loop
    Close #1
    j1 = txtj1.Text
    j2 = txtj2.Text
    j3 = txtj3.Text
    j4 = txtj4.Text
    j5 = txtj5.Text
    j6 = txtj6.Text
    PicResults.Print "Thanks for ur service with us!"
    PicResults.Print "*********************************************************************************"
    PicResults.Print "Quantity", "Price", "  Total"
    PicResults.Print "......................................................"


If j1 > 0 Then
total1 = j1 * price(1)
PicResults.Print j1; "(1)", price(1), "  "; total1
End If

If j2 > 0 Then
total2 = j2 * price(2)
PicResults.Print j2; "(2)", price(2), "  "; total2
End If

If j3 > 0 Then
total3 = j3 * price(3)
PicResults.Print j3; "(3)", price(3), "  "; total3
End If

If j4 > 0 Then
total4 = j4 * price(4)
PicResults.Print j4; "(4)", price(4), "  "; total4
End If

If j5 > 0 Then
total5 = j5 * price(5)
PicResults.Print j5; "(5)", price(5), "  "; total5
End If

If j6 > 0 Then
total6 = j6 * price(6)
PicResults.Print j6; "(6)", price(6), "  "; total6
End If

RunningTotal = total1 + total2 + total3 + total4 + total5 + total6
PicResults.Print
tax = RunningTotal * 0.07
PicResults.Print "Tax"; Tab(30); FormatCurrency(tax)
PicResults.Print "..............................................."
total = RunningTotal + tax
PicResults.Print "Total"; Tab(30); FormatCurrency(total)


    
   
End Sub

Private Sub cmdSubmit_Click()
Dim name As String
Dim address As String
name = InputBox("Insert User's Name", "Shipping and Handling")
address = InputBox("Insert Address to send equipment to", "Shipping and Handling")

MsgBox "The amount of " & FormatCurrency(total) & " will be billed to " & name & " at " & address & ". Your service is appreciated!", , "Shipping and Handling"

End Sub
