VERSION 5.00
Begin VB.Form frmDeliTacos 
   BackColor       =   &H000000FF&
   Caption         =   "Deli And Tacos"
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
      Height          =   1575
      Left            =   4800
      Picture         =   "frmDeliTacos.frx":0000
      ScaleHeight     =   1575
      ScaleWidth      =   2295
      TabIndex        =   23
      Top             =   120
      Width           =   2295
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
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdFocacciaSandwich 
      Caption         =   "Focaccia Sandwich"
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdBagelSandwich 
      Caption         =   "Bagel Sandwich"
      Height          =   1095
      Left            =   2280
      TabIndex        =   18
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdSoftShellTaco 
      Caption         =   "Soft Shell Taco"
      Height          =   1095
      Left            =   4440
      TabIndex        =   17
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdHardShellTaco 
      Caption         =   "Hard Shell Taco"
      Height          =   1095
      Left            =   6600
      TabIndex        =   16
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdNachos 
      Caption         =   "Nachos"
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdNachosSupreme 
      Caption         =   "Nachos Supreme"
      Height          =   1095
      Left            =   2280
      TabIndex        =   14
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdCupOfCheese 
      Caption         =   "Cup of Cheese"
      Height          =   1095
      Left            =   4440
      TabIndex        =   13
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdFreshWholeFruit 
      Caption         =   "Fresh Whole Fruit"
      Height          =   1095
      Left            =   6600
      TabIndex        =   12
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdThirdSub 
      Caption         =   "Third Sub"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Picture         =   "frmDeliTacos.frx":4643
      ScaleHeight     =   2535
      ScaleWidth      =   2655
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdWrap 
      Caption         =   "Wrap"
      Height          =   1095
      Left            =   6600
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdCroissantSandwich 
      Caption         =   "Croissant Sandwich"
      Height          =   1095
      Left            =   6600
      TabIndex        =   8
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdWholeVeggieSub 
      Caption         =   "Whole Veggie Sub"
      Height          =   1095
      Left            =   4440
      TabIndex        =   7
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdHalfVeggieSub 
      Caption         =   "Half Veggie Sub"
      Height          =   1095
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdThirdVeggieSub 
      Caption         =   "Third Veggie Sub"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdWholeSub 
      Caption         =   "Whole Sub"
      Height          =   1095
      Left            =   4440
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdHalfSub 
      Caption         =   "Half Sub"
      Height          =   1095
      Left            =   2280
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
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
      TabIndex        =   2
      Top             =   120
      Width           =   5415
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
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
      Left            =   120
      TabIndex        =   24
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Deli"
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
      Left            =   2880
      TabIndex        =   22
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmDeliTacos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sexton Dining Cash Register "\SextonDiningCashRegister.vpb"
'frmDeliTacos "\frmDeliTacos.frm"
'Paul Bivens
'March 22nd, 2006
'This form is used to ring up deli and taco items for purchase.

Option Explicit
Dim X As Integer
Dim Pos As Integer
Dim Y As Integer
Dim Found As Boolean
'Takes you back to the main form
Private Sub cmdBack_Click()
    frmMain.Show
    frmDeliTacos.Hide
End Sub
'Takes you to the pay form
Private Sub cmdPay_Click()
    frmPay.Show
    frmDeliTacos.Hide
End Sub
'Ends the program
Private Sub cmdQuit_Click()
    End
End Sub
'Takes you to the search and sort form
Private Sub cmdSearch_Click()
    frmSearch.Show
    frmDeliTacos.Hide
End Sub
'The following buttons are used to display the name and price of a particular item
'from within the name and price arrays.
'It does this by searching through the name array and finding a name that matches the
'text on the button.
Private Sub cmdThirdSub_Click()
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
        If cmdThirdSub.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdHalfSub_Click()
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
        If cmdHalfSub.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdWholeSub_Click()
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
        If cmdWholeSub.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdCroissantSandwich_Click()
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
        If cmdCroissantSandwich.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdWrap_Click()
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
        If cmdWrap.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdFocacciaSandwich_Click()
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
        If cmdFocacciaSandwich.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdBagelSandwich_Click()
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
        If cmdBagelSandwich.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdSoftShellTaco_Click()
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
        If cmdSoftShellTaco.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdHardShellTaco_Click()
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
        If cmdHardShellTaco.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdNachos_Click()
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
        If cmdNachos.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdNachosSupreme_Click()
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
        If cmdNachosSupreme.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdCupofCheese_Click()
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
        If cmdCupOfCheese.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdFreshWholeFruit_Click()
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
        If cmdFreshWholeFruit.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdThirdVeggieSub_Click()
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
        If cmdThirdVeggieSub.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdHalfVeggieSub_Click()
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
        If cmdHalfVeggieSub.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdWholeVeggieSub_Click()
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
        If cmdWholeVeggieSub.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub

