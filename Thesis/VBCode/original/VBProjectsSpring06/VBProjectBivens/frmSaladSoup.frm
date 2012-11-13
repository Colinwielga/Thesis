VERSION 5.00
Begin VB.Form frmSaladSoup 
   BackColor       =   &H000000FF&
   Caption         =   "Salads and Soups"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   6360
      Picture         =   "frmSaladSoup.frx":0000
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search and Sort"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   9600
      ScaleHeight     =   8355
      ScaleWidth      =   5355
      TabIndex        =   10
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton cmdFreshCutFruit 
      Caption         =   "Fresh Cut Fruit"
      Height          =   1095
      Left            =   2280
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdPastaSalad 
      Caption         =   "Pasta Salad"
      Height          =   1095
      Left            =   4440
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdChefSalad 
      Caption         =   "Chef Salad"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdBreadBowl 
      Caption         =   "Bread Bowl"
      Height          =   1095
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdDinnerSalad 
      Caption         =   "Dinner Salad"
      Height          =   1095
      Left            =   4440
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdVeggieBag 
      Caption         =   "Veggie Bag"
      Height          =   1095
      Left            =   6600
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdGardenSalad 
      Caption         =   "Garden Salad"
      Height          =   1095
      Left            =   6600
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Picture         =   "frmSaladSoup.frx":564E
      ScaleHeight     =   2535
      ScaleWidth      =   2655
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdCupOfSoup 
      Caption         =   "Cup Of Soup"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdYogurt 
      Caption         =   "Yogurt"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Paul Bivens"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7680
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Salad/Soup"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   39
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   15
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmSaladSoup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sexton Dining Cash Register "\SextonDiningCashRegister.vpb"
'frmSaladSoup "\frmSaladSoup.frm"
'Paul Bivens
'March 22nd, 2006
'This form is used to ring up salad and soup items for
'purchase.

Option Explicit
Dim X As Integer
Dim Pos As Integer
Dim Y As Integer
Dim Found As Boolean
'Takes you to the main form
Private Sub cmdBack_Click()
    frmMain.Show
    frmSaladSoup.Hide
End Sub
'The following buttons are used to display the name and price of a particular item
'from within the name and price arrays.
'It does this by searching through the name array and finding a name that matches the
'text on the button.
Private Sub cmdCupOfSoup_Click()
    Pos = 0
    Y = 0
    ArrayCounter = ArrayCounter + 1
    If ArrayCounter = 27 Then
        picOutput.Cls
        ArrayCounter = 0
        For X = 1 To 27
    End If
    
    Found = False
    Do While Found = False And Y < Size
        Pos = Pos + 1
        Y = Y + 1
        If cmdCupOfSoup.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdFreshCutFruit_Click()
    Pos = 0
    Y = 0
    ArrayCounter = ArrayCounter + 1
    If ArrayCounter = 27 Then
        picOutput.Cls
        ArrayCounter = 0
    End If
    Found = False
    Do While Found = False And Y < Size
        Pos = Pos + 1
        Y = Y + 1
        If cmdFreshCutFruit.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdPastaSalad_Click()
    Pos = 0
    Y = 0
    ArrayCounter = ArrayCounter + 1
    If ArrayCounter = 27 Then
        picOutput.Cls
        ArrayCounter = 0
    End If
    Found = False
    Do While Found = False And Y < Size
        Pos = Pos + 1
        Y = Y + 1
        If cmdPastaSalad.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdGardenSalad_Click()
    Pos = 0
    Y = 0
    ArrayCounter = ArrayCounter + 1
    If ArrayCounter = 27 Then
        picOutput.Cls
        ArrayCounter = 0
    End If
    Found = False
    Do While Found = False And Y < Size
        Pos = Pos + 1
        Y = Y + 1
        If cmdGardenSalad.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdChefSalad_Click()
    Pos = 0
    Y = 0
    ArrayCounter = ArrayCounter + 1
    If ArrayCounter = 27 Then
        picOutput.Cls
        ArrayCounter = 0
    End If
    Found = False
    Do While Found = False And Y < Size
        Pos = Pos + 1
        Y = Y + 1
        If cmdChefSalad.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdBreadBowl_Click()
    Pos = 0
    Y = 0
    ArrayCounter = ArrayCounter + 1
    If ArrayCounter = 27 Then
        picOutput.Cls
        ArrayCounter = 0
    End If
    Found = False
    Do While Found = False And Y < Size
        Pos = Pos + 1
        Y = Y + 1
        If cmdBreadBowl.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdDinnerSalad_Click()
    Pos = 0
    Y = 0
    ArrayCounter = ArrayCounter + 1
    If ArrayCounter = 27 Then
        picOutput.Cls
        ArrayCounter = 0
    End If
    Found = False
    Do While Found = False And Y < Size
        Pos = Pos + 1
        Y = Y + 1
        If cmdDinnerSalad.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
'Takes you to the pay form.
Private Sub cmdPay_Click()
    frmPay.Show
    frmSaladSoup.Hide
End Sub
'Ends the program.
Private Sub cmdQuit_Click()
    End
End Sub
'Takes you to the search and sort form.
Private Sub cmdSearch_Click()
    frmSearch.Show
    frmSaladSoup.Hide
End Sub

Private Sub cmdVeggieBag_Click()
    Pos = 0
    Y = 0
    ArrayCounter = ArrayCounter + 1
    If ArrayCounter = 27 Then
        picOutput.Cls
        ArrayCounter = 0
    End If
    Found = False
    Do While Found = False And Y < Size
        Pos = Pos + 1
        Y = Y + 1
        If cmdVeggieBag.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdYogurt_Click()
    Pos = 0
    Y = 0
    ArrayCounter = ArrayCounter + 1
    If ArrayCounter = 27 Then
        picOutput.Cls
        ArrayCounter = 0
    End If
    Found = False
    Do While Found = False And Y < Size
        Pos = Pos + 1
        Y = Y + 1
        If cmdYogurt.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub

