VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H000000FF&
   Caption         =   "Search"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   735
      Left            =   11640
      TabIndex        =   12
      Top             =   120
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
      Height          =   735
      Left            =   11640
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
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
      Height          =   735
      Left            =   11640
      TabIndex        =   10
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortPrice 
      Caption         =   "Sort By Price"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11640
      TabIndex        =   9
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmdSortName 
      Caption         =   "Sort By Name"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11640
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display All Items"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   2535
   End
   Begin VB.PictureBox picOutput2 
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
      Height          =   10935
      Left            =   7200
      ScaleHeight     =   10875
      ScaleWidth      =   4275
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdSearchHigher 
      Caption         =   "Find Items That Cost More Than:"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearchLower 
      Caption         =   "Find Items That Cost Less Than:"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   2535
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
      Height          =   10935
      Left            =   2880
      ScaleHeight     =   10875
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdSearchName 
      Caption         =   "Search By Name"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   7560
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearchPrice 
      Caption         =   "Search By Price"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Picture         =   "frmSearch.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   120
      Width           =   2655
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
      Left            =   13440
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sexton Dining Cash Register "\SextonDiningCashRegister.vpb"
'frmSearch "\frmSearch.frm"
'Paul Bivens
'March 22nd, 2006
'This form is used to search for item names and prices and
'sort that information.


Option Explicit
Dim Price As Single
Dim Namee As String
Dim Pos As Integer
Dim X As Integer
Dim Y As Integer
Dim Found As Boolean
Dim TempNameArray(1 To 1000) As String
Dim TempPriceArray(1 To 1000) As Single
Dim TempName As String
Dim TempPrice As Single
Dim TempSize As Integer
Dim Z As Integer
Dim Pass As Integer
Dim A As String
'Takes you back to your main form.
Private Sub cmdBack_Click()
    frmMain.Show
    frmSearch.Hide
End Sub
'This button is used to display all of the items that are within the array. The only
'real complication with this was that there was no pic box big enough to hold all of
'the values. I fixed this by setting up a second pic box and having the program begin
'printing there when the Y value, which was simply a counter, reached 36.
'It also inputs the values into an array used later for sorting.
Private Sub cmdDisplay_Click()
    picOutput.Cls
    Pos = 0
    Y = 0
    For X = 1 To Size
        Pos = Pos + 1
        If Y <= 35 Then
            Y = Y + 1
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
        ElseIf Y > 35 Then
            Y = Y + 1
            picOutput2.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
        End If
    Next X
    TempSize = Y
End Sub
'Displays the pay form.
Private Sub cmdPay_Click()
    frmPay.Show
    frmSearch.Hide
End Sub
'Ends the program.
Private Sub cmdQuit_Click()
    End
End Sub
'This button has the user input a numerical value with which it then searches for all
'values higher than.  It will then display each of these values in the order they are
'found and this also displays the information in two boxes if necessary like the display
'button did. It also puts the found values into arrays for sorting.
Private Sub cmdSearchHigher_Click()
    Y = 0
    picOutput.Cls
    picOutput2.Cls
    Price = InputBox("Input Price to Search For", "Price Search")
    Pos = 0
    For X = 1 To Size
        Pos = Pos + 1
        If Price < priceArray(Pos) And Y <= 35 Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Y = Y + 1
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
        ElseIf Price < priceArray(Pos) And Y > 35 Then
            Y = Y + 1
            picOutput2.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
        End If
    Next X
    TempSize = Y
End Sub
'This button has the user input a numerical value with which it then searches for all
'values lower than.  It will then display each of these values in the order they are
'found and this also displays the information in two boxes if necessary like the
'previous buttons did. It also puts the found values into arrays for sorting.
Private Sub cmdSearchLower_Click()
    Y = 0
    picOutput.Cls
    picOutput2.Cls
    Price = InputBox("Input Price to Search For", "Price Search")
    Pos = 0
    For X = 1 To Size
        Pos = Pos + 1
        If Price > priceArray(Pos) And Y <= 35 Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Y = Y + 1
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
        
        ElseIf Price > priceArray(Pos) And Y > 35 Then
            Y = Y + 1
            picOutput2.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
        End If
    Next X
    TempSize = Y
End Sub
'This button has the user input a name with which it then searches for all
'values with the same name.  It will then display each of these values in the order
'they are found and this also displays the information in two boxes if necessary like
'the previous buttons did. It also puts the found values into arrays for sorting.
Private Sub cmdSearchName_Click()
    picOutput.Cls
    picOutput2.Cls
    Found = False
    Y = 0
    A = InputBox("Input Item For Search", "Item Search")
    For Pos = 1 To Size
        If InStr(nameArray(Pos), A) And Y <= 35 Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Y = Y + 1
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
            Found = True
        ElseIf InStr(nameArray(Pos), A) And Y > 35 Then
            picOutput2.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Y = Y + 1
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
        End If
    Next Pos
    If Found = False Then
        MsgBox "Consider checking your spelling or captitalization.", , "Item Not Found"
    End If
    TempSize = Y
End Sub
'This button has the user input a numerical value with which it then searches for all
'values equal to that number.  It will then display each of these values in the order
'they are found and this also displays the information in two boxes if necessary like
'the display button did. It also puts the found values into arrays for sorting.
Private Sub cmdSearchPrice_Click()
    Y = 0
    Found = False
    picOutput.Cls
    picOutput2.Cls
    Price = InputBox("Input Price to Search For", "Price Search")
    Pos = 0
    For X = 1 To Size
        Pos = Pos + 1
        If Price = priceArray(Pos) And Y <= 35 Then
            picOutput.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            Y = Y + 1
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
            Found = True
        ElseIf Price = priceArray(Pos) And Y > 35 Then
            Y = Y + 1
            picOutput2.Print nameArray(Pos); Tab(30); FormatCurrency(priceArray(Pos))
            TempNameArray(Y) = nameArray(Pos)
            TempPriceArray(Y) = priceArray(Pos)
        End If
    Next X
    If Found = False Then
        MsgBox "There are no items at that price.", , "Input Invalid"
    End If
    TempSize = Y
End Sub
'This button takes the information from the array in which all the found values above
'were found in. It sorts those arrays in alphabetical order and sorts the corresponding
'price array.  It then displays both of these in the pic box.
Private Sub cmdSortName_Click()
    picOutput.Cls
    picOutput2.Cls
    Z = 0
    Pos = 0
    For Pass = 1 To TempSize - 1
        Pos = 0
        For X = 1 To TempSize - Pass
            Pos = Pos + 1
            If TempNameArray(Pos) > TempNameArray(Pos + 1) Then
                TempName = TempNameArray(Pos)
                TempNameArray(Pos) = TempNameArray(Pos + 1)
                TempNameArray(Pos + 1) = TempName
                TempPrice = TempPriceArray(Pos)
                TempPriceArray(Pos) = TempPriceArray(Pos + 1)
                TempPriceArray(Pos + 1) = TempPrice
            End If
        Next X
    Next Pass
    For X = 1 To TempSize
        Z = Z + 1
        If Z <= 36 Then
            picOutput.Print TempNameArray(X); Tab(30); FormatCurrency(TempPriceArray(X))
        End If
        If Z > 36 Then
            picOutput2.Print TempNameArray(X); Tab(30); FormatCurrency(TempPriceArray(X))
        End If
    Next X
End Sub
'This button does the same as the name sorting button but sorts the prices instead.
Private Sub cmdSortPrice_Click()
    picOutput.Cls
    picOutput2.Cls
    Z = 0
    Pos = 0
    For Pass = 1 To TempSize - 1
        Pos = 0
        For X = 1 To TempSize - Pass
            Pos = Pos + 1
            If TempPriceArray(Pos) > TempPriceArray(Pos + 1) Then
                TempName = TempNameArray(Pos)
                TempNameArray(Pos) = TempNameArray(Pos + 1)
                TempNameArray(Pos + 1) = TempName
                TempPrice = TempPriceArray(Pos)
                TempPriceArray(Pos) = TempPriceArray(Pos + 1)
                TempPriceArray(Pos + 1) = TempPrice
            End If
        Next X
    Next Pass
    For X = 1 To TempSize
        Z = Z + 1
        If Z <= 36 Then
            picOutput.Print TempNameArray(X); Tab(30); FormatCurrency(TempPriceArray(X))
        End If
        If Z > 36 Then
            picOutput2.Print TempNameArray(X); Tab(30); FormatCurrency(TempPriceArray(X))
        End If
    Next X
End Sub
