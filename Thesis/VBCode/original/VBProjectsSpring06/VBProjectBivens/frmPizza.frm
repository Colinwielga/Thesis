VERSION 5.00
Begin VB.Form frmPizza 
   BackColor       =   &H000000FF&
   Caption         =   "Pizza"
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
      Height          =   1695
      Left            =   4680
      Picture         =   "frmPizza.frx":0000
      ScaleHeight     =   1695
      ScaleWidth      =   1935
      TabIndex        =   12
      Top             =   0
      Width           =   1935
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSliceOfPizza 
      Caption         =   "Slice of Pizza"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Picture         =   "frmPizza.frx":4CAA
      ScaleHeight     =   2535
      ScaleWidth      =   2655
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdPretzelAndCheese 
      Caption         =   "Pretzel And Cheese"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalzone 
      Caption         =   "Calzone"
      Height          =   1095
      Left            =   6600
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdExtraTopping 
      Caption         =   "Extra Topping"
      Height          =   1095
      Left            =   4440
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdWholePizza 
      Caption         =   "Whole Pizza"
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
      Left            =   7680
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Pizza"
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
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmPizza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sexton Dining Cash Register "\SextonDiningCashRegister.vpb"
'frmPizza "\frmPizza.frm"
'Paul Bivens
'March 22nd, 2006
'This form is used to ring up pizza items for purchase.

Option Explicit
Dim X As Integer
Dim Pos As Integer
Dim Y As Integer
Dim Found As Boolean
'Returns to the main form
Private Sub cmdBack_Click()
    frmMain.Show
    frmPizza.Hide
End Sub
'Takes you to the pay form
Private Sub cmdPay_Click()
    frmPay.Show
    frmPizza.Hide
End Sub
'Ends the program
Private Sub cmdQuit_Click()
    End
End Sub
'Takes you to the search and sort form
Private Sub cmdSearch_Click()
    frmSearch.Show
    frmPizza.Hide
End Sub
'The following buttons are used to display the name and price of a particular item
'from within the name and price arrays.
'It does this by searching through the name array and finding a name that matches the
'text on the button.
Private Sub cmdSliceOfPizza_Click()
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
        If cmdSliceOfPizza.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdWholePizza_Click()
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
        If cmdWholePizza.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdExtraTopping_Click()
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
        If cmdExtraTopping.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdCalzone_Click()
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
        If cmdCalzone.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub
Private Sub cmdPretzelAndCheese_Click()
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
        If cmdPretzelAndCheese.Caption = nameArray(Pos) Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Found = True
            Sum = Sum + priceArray(Pos)
        End If
    Loop
    If Found = False Then
        MsgBox "Error"
    End If
End Sub

