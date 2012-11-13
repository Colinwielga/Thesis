VERSION 5.00
Begin VB.Form Home 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   Picture         =   "Home.frx":0000
   ScaleHeight     =   5490
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3240
      TabIndex        =   21
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3240
      TabIndex        =   20
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0000FFFF&
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000FFFF&
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Total"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4680
      Width           =   3495
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   7440
      ScaleHeight     =   4995
      ScaleWidth      =   3075
      TabIndex        =   15
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000080FF&
      Caption         =   "Other"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H000080FF&
      Caption         =   "Bakery / Salad / Soup"
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000080FF&
      Caption         =   "Snacks"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000080FF&
      Caption         =   "Drinks"
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000080FF&
      Caption         =   "Grill"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H000080FF&
      Caption         =   "Deli"
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000C0&
      Caption         =   "Credit Card (with tax)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000C0&
      Caption         =   "Cash (with tax)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000C0&
      Caption         =   "Charge (with tax)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "Flex (no tax)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4440
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "Punch (no tax)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4440
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H000080FF&
      Caption         =   "   Quantity"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Caption         =   "   Quantity"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   2640
      Width           =   855
   End
   Begin VB.Line Line5 
      X1              =   4920
      X2              =   6480
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line4 
      X1              =   1560
      X2              =   2760
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   120
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      X1              =   7200
      X2              =   7200
      Y1              =   240
      Y2              =   5400
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dish Options"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Options"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4200
      Y1              =   1320
      Y2              =   5400
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the dishes you have chosen. After you have selected the dishes, please choose a payment option."
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Sexton Dining."
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Flex As Single, Subtotal As Single, over As Single, Punch As Single, Charge As Single, Cash As Single, Credit As Single, Tax As Single, GrandTotal As Single, FinalTotal As Single, Total1 As Single, Total2 As Single, Total3 As Single, Total4 As Single, Total5 As Single, Total6 As Single, temptotal As Single, Pass As Single, Name1 As String, Name2 As String, Name3 As String, Name4 As String, Name5 As String, Name6 As String, tempname As String

Private Sub Command1_Click()
Picture1.Print ""
Punch = 4.4
FinalTotal = FinalTotal - Punch
If FinalTotal = 0 Then
    Picture1.Print "Thank you! Have a nice day."
    ElseIf FinalTotal > 0 Then
    Picture1.Print "You still owe "; FormatCurrency(FinalTotal)
    ElseIf FinalTotal < 0 Then
    over = (0 - FinalTotal)
    Picture1.Print "Thank you! Have a nice day."
End If
End Sub

Private Sub Command10_Click()
BSS.Show
Home.Hide
End Sub

Private Sub Command11_Click()
Dim Otheritem As String, Othercost As Single, Otherq As Integer, Other As Single
Otheritem = InputBox("What item are you buying?")
Othercost = InputBox("How much does one of those items cost? I.E. If it's $.50 enter .5, if it's $1.75 enter 1.75")
Otherq = Text2.Text
Other = Othercost * Otherq
answer = InputBox("Do you want to checkout? If yes you can't get anymore other items. If yes enter 1, if no enter 0?")
If answer = 1 Then
    OtherTotal = OtherTotal + Other
    Command11.Enabled = False
    Command13.Enabled = True
    ElseIf answer = 0 Then
        OtherTotal = OtherTotal + Other
        Other = 0
        Otherq = 0
        Othercost = 0
Else
End If
End Sub

Private Sub Command12_Click()
End
End Sub

Private Sub Command13_Click()
FinalTotal = GrillTotal + DrinkTotal + SnackTotal + DeliTotal + BakeryTotal + OtherTotal
Tax = FinalTotal * 0.065
GrandTotal = Tax + FinalTotal
Tax = FormatNumber(Tax)
GrandTotal = FormatNumber(GrandTotal)
FinalTotal = FormatNumber(FinalTotal)
Picture1.Print "Grill Total", "= "; FormatCurrency(GrillTotal)
Picture1.Print "Drink Total", "= "; FormatCurrency(DrinkTotal)
Picture1.Print "Deli Total", "= "; FormatCurrency(DeliTotal)
Picture1.Print "B.S.S. Total", "= "; FormatCurrency(BakeryTotal)
Picture1.Print "Snack Total", "= "; FormatCurrency(SnackTotal)
Picture1.Print "Other Total", "= "; FormatCurrency(OtherTotal)
Picture1.Print "****************************"
Picture1.Print "Subtotal", "= "; FormatCurrency(FinalTotal)
Picture1.Print "****************************"
Picture1.Print "Tax", "= "; FormatCurrency(Tax)
Picture1.Print "****************************"
Picture1.Print "Grand Total", "= "; FormatCurrency(GrandTotal)
Total1 = GrillTotal
Total2 = DrinkTotal
Total3 = DeliTotal
Total4 = BakeryTotal
Total5 = SnackTotal
Total6 = OtherTotal
Name1 = "Grill"
Name2 = "Drink"
Name3 = "Deli"
Name4 = "Bakery / Salad / Soup"
Name5 = "Snack"
Name6 = "Other"
For Pass = 1 To 6 - 1
    If Total1 < Total2 Then
        temptotal = Total1
        Total1 = Total2
        Total2 = temptotal
        tempname = Name1
        Name1 = Name2
        Name2 = tempname
    End If
    If Total2 < Total3 Then
        temptotal = Total2
        Total2 = Total3
        Total3 = temptotal
        tempname = Name2
        Name2 = Name3
        Name3 = tempname
    End If
    If Total3 < Total4 Then
        temptotal = Total3
        Total3 = Total4
        Total4 = temptotal
        tempname = Name3
        Name3 = Name4
        Name4 = tempname
    End If
    If Total4 < Total5 Then
        temptotal = Total4
        Total4 = Total5
        Total5 = temptotal
        tempname = Name4
        Name4 = Name5
        Name5 = tempname
    End If
    If Total5 < Total6 Then
        temptotal = Total5
        Total5 = Total6
        Total6 = temptotal
        tempname = Name5
        Name5 = Name6
        Name6 = tempname
    End If
Next Pass
MsgBox ("You paid the most in " & Name1 & " with a total of" & FormatCurrency(Total1))
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
Command13.Enabled = False
End Sub

Private Sub Command14_Click()
Picture1.Cls
Total = 0
Subtotal = 0
Tax = 0
Command13.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
End Sub

Private Sub Command15_Click()
Picture1.Print "Item", "Subtotal"
Picture1.Print "****************************"
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
Command15.Enabled = False
End Sub

Private Sub Command2_Click()
Picture1.Print ""
Flex = InputBox("How much would you like to flex?")

FinalTotal = FinalTotal - Flex
If FinalTotal = 0 Then
    Picture1.Print "Thank you! Have a nice day."
    ElseIf FinalTotal > 0 Then
    Picture1.Print "You still owe "; FormatCurrency(FinalTotal)
    ElseIf FinalTotal < 0 Then
    over = (0 - FinalTotal)
    Picture1.Print "You paid over the amount,"
    Picture1.Print " you get "; FormatCurrency(over); " back"
End If
End Sub

Private Sub Command3_Click()
Picture1.Print ""
Charge = InputBox("How much would you like to charge?")
GrandTotal = GrandTotal - Charge
If GrandTotal = 0 Then
    Picture1.Print "Thank you! Have a nice day."
    ElseIf GrandTotal > 0 Then
    Picture1.Print "You still owe "; FormatCurrency(GrandTotal)
    ElseIf GrandTotal < 0 Then
    over = (0 - GrandTotal)
    Picture1.Print "You paid over the amount,"
    Picture1.Print " you get "; FormatCurrency(over); " back"
End If
End Sub

Private Sub Command4_Click()
Picture1.Print " "
Cash = InputBox("How much cash would you like to use?")
GrandTotal = GrandTotal - Cash
If GrandTotal = 0 Then
    Picture1.Print "Thank you! Have a nice day."
    ElseIf GrandTotal > 0 Then
    Picture1.Print "You still owe "; FormatCurrency(GrandTotal)
    ElseIf GrandTotal < 0 Then
    over = (0 - GrandTotal)
    Picture1.Print "You paid over the amount,"
    Picture1.Print " you get "; FormatCurrency(over); " back"
End If
End Sub

Private Sub Command5_Click()
Picture1.Print " "
Credit = InputBox("How much would you like to charge?")
GrandTotal = GrandTotal - Credit
If GrandTotal = 0 Then
    Picture1.Print " Thank you! Have a nice day."
    ElseIf GrandTotal > 0 Then
    Picture1.Print "You still owe "; FormatCurrency(GrandTotal)
    ElseIf GrandTotal < 0 Then
    over = (0 - GrandTotal)
    Picture1.Print "You paid over the amount,"
    Picture1.Print " you get "; FormatCurrency(over); " back"
End If
End Sub

Private Sub Command6_Click()
Deli.Show
Home.Hide
End Sub

Private Sub Command7_Click()
Home.Hide
Grill.Show
End Sub

Private Sub Command8_Click()
Drinks.Show
Home.Hide
End Sub

Private Sub Command9_Click()
Dim Snackitem As String, Snackcost As Single, Snackq As Integer, Snack As Single
Dim answer As Single
Snackitem = InputBox("What item are you buying?")
Snackcost = InputBox("How much does one of those items cost? I.E. If it's $.50 enter .5, if it's $1.75 enter 1.75")
Snackq = Text1.Text
Snack = Snackcost * Snackq
answer = InputBox("Do you want to checkout? If yes you can't get anymore snack items. If yes enter 1, if no enter 0?")
If answer = 1 Then
    SnackTotal = SnackTotal + Snack
    Command9.Enabled = False
    Command13.Enabled = True
    ElseIf answer = 0 Then
        SnackTotal = SnackTotal + Snack
        Snack = 0
        Snackq = 0
        Snackcost = 0
Else
End If
End Sub

