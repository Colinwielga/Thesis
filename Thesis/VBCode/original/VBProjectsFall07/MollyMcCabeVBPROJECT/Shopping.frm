VERSION 5.00
Begin VB.Form frmShopping 
   BackColor       =   &H00400000&
   Caption         =   "Shopping"
   ClientHeight    =   8595
   ClientLeft      =   2085
   ClientTop       =   1020
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10710
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit Twins Territory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000C0&
      Caption         =   "Back to Twins Territory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H000000C0&
      Caption         =   "Done Shopping"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H000000C0&
      Caption         =   "Start Shopping!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   5640
      ScaleHeight     =   5355
      ScaleWidth      =   4755
      TabIndex        =   10
      Top             =   720
      Width           =   4815
   End
   Begin VB.CommandButton cmdWHoodie 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   240
      Picture         =   "Shopping.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdSleepSet 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   5640
      Picture         =   "Shopping.frx":7422
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdMHat 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   240
      Picture         =   "Shopping.frx":E6A7
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdWHat 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   2040
      Picture         =   "Shopping.frx":15153
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdKHat 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   3840
      Picture         =   "Shopping.frx":1BAC9
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdTshirt 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   3840
      Picture         =   "Shopping.frx":22909
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdWindbreaker 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   2040
      Picture         =   "Shopping.frx":29364
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdWJersey 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   3840
      Picture         =   "Shopping.frx":313D5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdBJersey 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   2040
      Picture         =   "Shopping.frx":3A99B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdSJersey 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   240
      Picture         =   "Shopping.frx":42406
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblURL 
      BackColor       =   &H000000C0&
      Caption         =   "http://shop.mlb.com/shop/index.jsp?categoryId=1452357"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   36
      Top             =   6840
      Width           =   3135
   End
   Begin VB.Label lblMore 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "For More Twins Apparel Go To:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   31
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label lblPrice10 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$22.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   30
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label lblPrice9 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$19.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   29
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label lblPrice8 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$49.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   28
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label lblPrice7 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$54.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label lblSleepSet 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Toddler Sleep Set"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label lblTshirt 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "T-Shirt"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label lblMJacket 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Men's Jacket"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   24
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label lblWHoodie 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Women's Hoodie"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label lblPrice4 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$27.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblPrice5 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$18.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblPrice6 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$16.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblKCap 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Kid's Cap"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblWCap 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Women's Cap"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblMCap 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Men's Cap"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblPrice3 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$79.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblPrice2 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$219.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblWJersey 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Customized Jersey (Woman's)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   14
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblBJersey 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Customized Jersey (Blue)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   13
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblPrice1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "$219.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblSJersey 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Customized Jersey (Stripes)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "frmShopping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SubTotal As Single
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub lblURL_DragDrop(Source As Control, X As Single, Y As Single)
    If Source Is lblURL Then
        With lblURL
            .Font.Underline = False
            .ForeColor = vbBlack
            ' Call ShellExecute(0&, vbNullString, "Mailto:" & .Caption, vbNullString, vbNullString, vbNormalFocus)
            Call ShellExecute(0&, vbNullString, .Caption, vbNullString, vbNullString, vbNormalFocus)
        End With
    End If
End Sub

Private Sub lblURL_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbLeave Then
        With lblURL
            .Drag vbEndDrag
            .Font.Underline = False
            .ForeColor = vbBlack
        End With
    End If
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblURL
        .ForeColor = vbBlue
        .Font.Underline = True
        .Drag vbBeginDrag
    End With
End Sub

Private Sub cmdBack_Click()
    frmShopping.Hide 'hides shopping form
    frmMain.Show 'shows Main form
End Sub

Private Sub cmdBJersey_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 219.99
size = InputBox("Enter a size S, M, L, or XL", "Size")
'If the size entered is valid, print item size and price, and add price to subtotal
If (size <> "S") And (size <> "M") And (size <> "L") And (size <> "XL") Then
    MsgBox "You Must Enter A Valid Size", , "Error"
Else
    SubTotal = SubTotal + price
    picResults.Print "Customized Jersey (Blue)"; Tab(33); size; Tab(50); FormatCurrency(price)
End If
End Sub

Private Sub cmdDone_Click()
'declare variable
Dim Tax As Single
'initialize Tax as 7% of the Subtotal
Tax = 0.07
'print subtotal, tax, and total of purchase
picResults.Print "*************************************************************************************"
picResults.Print "SubTotal:", FormatCurrency(SubTotal)
picResults.Print "Tax:", FormatCurrency(SubTotal * Tax)
picResults.Print "Total Cost:", FormatCurrency(SubTotal + (SubTotal * Tax))

'disable item buttons
    cmdSJersey.Enabled = False
    cmdBJersey.Enabled = False
    cmdWJersey.Enabled = False
    cmdMHat.Enabled = False
    cmdWHat.Enabled = False
    cmdKHat.Enabled = False
    cmdWHoodie.Enabled = False
    cmdWindbreaker.Enabled = False
    cmdTshirt.Enabled = False
    cmdSleepSet.Enabled = False
    
'enable start button
    cmdStart.Enabled = True
    
'disable done button
    cmdDone.Enabled = False
End Sub

Private Sub cmdExit_Click()
    End 'end progam
End Sub

Private Sub cmdKHat_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 16.99
size = "OneSizeFitsAll"
'print item size and price, and add price to subtotal
    SubTotal = SubTotal + price
    picResults.Print "Kid's Cap"; Tab(33); size; Tab(50); FormatCurrency(price)
End Sub

Private Sub cmdMHat_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 27.99
size = "OneSizeFitsAll"
'print item size and price, and add price to subtotal
    SubTotal = SubTotal + price
    picResults.Print "Men's Cap"; Tab(33); size; Tab(50); FormatCurrency(price)
End Sub

Private Sub cmdSJersey_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 219.99
size = InputBox("Enter a size S, M, L, or XL", "Size")
'If the size entered is valid, print item size and price, and add price to subtotal
If (size <> "S") And (size <> "M") And (size <> "L") And (size <> "XL") Then
    MsgBox "You Must Enter A Valid Size", , "Error"
Else
    SubTotal = SubTotal + price
    picResults.Print "Customized Jersey (Stripes)"; Tab(33); size; Tab(50); FormatCurrency(price)
End If

End Sub


Private Sub cmdSleepSet_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 22.99
size = InputBox("Enter a size S, M, L, or XL", "Size")
'If the size entered is valid, print item size and price, and add price to subtotal
If (size <> "S") And (size <> "M") And (size <> "L") And (size <> "XL") Then
    MsgBox "You Must Enter A Valid Size", , "Error"
Else
    SubTotal = SubTotal + price
    picResults.Print "Toddler Sleep Set"; Tab(33); size; Tab(50); FormatCurrency(price)
End If
End Sub

Private Sub cmdStart_Click()
'clear picture box
    picResults.Cls
'enable Item buttons
    cmdSJersey.Enabled = True
    cmdBJersey.Enabled = True
    cmdWJersey.Enabled = True
    cmdMHat.Enabled = True
    cmdWHat.Enabled = True
    cmdKHat.Enabled = True
    cmdWHoodie.Enabled = True
    cmdWindbreaker.Enabled = True
    cmdTshirt.Enabled = True
    cmdSleepSet.Enabled = True
    cmdDone.Enabled = True
    
'set Subtotal to 0
SubTotal = 0

'print top of reciept
    picResults.Print "Total Purchases"
    picResults.Print "Item"; Tab(33); "Size"; Tab(50); "Cost"
    picResults.Print "************************************************************************************************"
End Sub

Private Sub cmdTshirt_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 19.99
size = InputBox("Enter a size S, M, L, or XL", "Size")
'If the size entered is valid, print item size and price, and add price to subtotal
If (size <> "S") And (size <> "M") And (size <> "L") And (size <> "XL") Then
    MsgBox "You Must Enter A Valid Size", , "Error"
Else
    SubTotal = SubTotal + price
    picResults.Print "T-Shirt"; Tab(33); size; Tab(50); FormatCurrency(price)
End If
End Sub

Private Sub cmdWHat_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 18.99
size = "OneSizeFitsAll"
'print item size and price, and add price to subtotal
    SubTotal = SubTotal + price
    picResults.Print "Women's Cap"; Tab(33); size; Tab(50); FormatCurrency(price)
End Sub

Private Sub cmdWHoodie_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 54.99
size = InputBox("Enter a size S, M, L, or XL", "Size")
'If the size entered is valid, print item size and price, and add price to subtotal
If (size <> "S") And (size <> "M") And (size <> "L") And (size <> "XL") Then
    MsgBox "You Must Enter A Valid Size", , "Error"
Else
    SubTotal = SubTotal + price
    picResults.Print "Women's Hoodie"; Tab(33); size; Tab(50); FormatCurrency(price)
End If
End Sub

Private Sub cmdWindbreaker_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 49.99
size = InputBox("Enter a size S, M, L, or XL", "Size")
'If the size entered is valid, print item size and price, and add price to subtotal
If (size <> "S") And (size <> "M") And (size <> "L") And (size <> "XL") Then
    MsgBox "You Must Enter A Valid Size", , "Error"
Else
    SubTotal = SubTotal + price
    picResults.Print "Men's Jacket"; Tab(33); size; Tab(50); FormatCurrency(price)
End If
End Sub

Private Sub cmdWJersey_Click()
'disable start button
    cmdStart.Enabled = False
'declare variables
Dim size As String
Dim price As Single
'initialize variables
price = 79.99
size = InputBox("Enter a size S, M, L, or XL", "Size")
'If the size entered is valid, print item size and price, and add price to subtotal
If (size <> "S") And (size <> "M") And (size <> "L") And (size <> "XL") Then
    MsgBox "You Must Enter A Valid Size", , "Error"
Else
    SubTotal = SubTotal + price
    picResults.Print "Customized Jersey (Women's)"; Tab(33); size; Tab(50); FormatCurrency(price)
End If
End Sub


