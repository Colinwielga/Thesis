VERSION 5.00
Begin VB.Form frmCheckOut 
   BackColor       =   &H00000000&
   Caption         =   "Check Out"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form7"
   ScaleHeight     =   6900
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      Height          =   4455
      Left            =   2280
      ScaleHeight     =   4395
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   1440
      Width           =   6255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label lblCheckout 
      BackColor       =   &H00000000&
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10080
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmCheckOut.Hide
frmHome.Show
'move to Home form
End Sub


Private Sub cmdConfirm_Click()
picOutput.Cls

If buns > 0 Then
    costbuns = buns * 3.25
    picOutput.Print buns; "bag(s) of buns @ $3.25 ="; Tab(38); FormatCurrency(costbuns)
End If

If bread > 0 Then
    costbread = bread * 2.5
    picOutput.Print bread; "loaf(s) of bread @ $2.50 ="; Tab(38); FormatCurrency(costbread)
End If

If pie > 0 Then
    costpie = pie * 12
    picOutput.Print pie; "pie(s) @ $12 ="; Tab(38); FormatCurrency(costpie)
End If

If pizza > 0 Then
    costpizza = pizza * 3.79
    picOutput.Print pizza; "pizza(s) @ $3.79 ="; Tab(38); FormatCurrency(costpizza)
End If

If shrimp > 0 Then
    costshrimp = shrimp * 4.69
    picOutput.Print shrimp; "bag(s) of shrimp @ $4.69 ="; Tab(38); FormatCurrency(costshrimp)
End If

If steak > 0 Then
    costshrimp = shrimp * 4.69
    picOutput.Print shrimp; "bag(s) of shrimp @ $4.69 ="; Tab(38); FormatCurrency(costshrimp)
End If

If deli > 0 Then
    costdeli = deli * 56
    picOutput.Print deli; "platters of deli meat @ $56 ="; Tab(38); FormatCurrency(costdeli)
End If

If cheese > 0 Then
    costcheese = cheese * 43
    picOutput.Print cheese; "platters of cheese @ $43 ="; Tab(38); FormatCurrency(costcheese)
End If

If salmon > 0 Then
    costsalmon = salmon * 13.99
    picOutput.Print buns; "lb(s) of salmon @ $13.99/lb ="; Tab(38); FormatCurrency(costsalmon)
End If

If apples > 0 Then
    costapples = apples * 5.45
    picOutput.Print apples; "bag(s) of apples @ $5.45 ="; Tab(38); FormatCurrency(costapples)
End If

If oranges > 0 Then
    costoranges = oranges * 5.95
    picOutput.Print oranges; "bag(s) of oranges @ $5.95 ="; Tab(38); FormatCurrency(costoranges)
End If

If bananas > 0 Then
    costbananas = bananas * 0.99
    picOutput.Print bananas; "lb(s) of bananas @ $.99/lb ="; Tab(38); FormatCurrency(costbananas)
End If

If carrots > 0 Then
    costcarrots = carrots * 0.74
    picOutput.Print carrots; "lb(s) of carrots @ $.74/lb ="; Tab(38); FormatCurrency(costcarrots)
End If

If tomatoes > 0 Then
    costtomatoes = tomatoes * 0.98
    picOutput.Print tomatoes; "lb(s) of tomatoes @ $.98/lb ="; Tab(38); FormatCurrency(costtomatoes)
End If

If lettuce > 0 Then
    costlettuce = lettuce * 0.76
    picOutput.Print lettuce; "lb(s) of lettuce @ $.76/lb ="; Tab(38); FormatCurrency(costlettuce)
End If

If toothpaste > 0 Then
    costtoothpaste = toothpaste * 2.19
    picOutput.Print toothpaste; "tube(s) of toothpaste @ $2.19 ="; Tab(38); FormatCurrency(costtoothpaste)
End If

If shampoo > 0 Then
    costshampoo = shampoo * 1.89
    picOutput.Print shampoo; "container(s) of shampoo @ $1.89 ="; Tab(38); FormatCurrency(costshampoo)
End If

If floss > 0 Then
    costfloss = floss * 0.49
    picOutput.Print floss; "containter(s) of floss @ $.49 ="; Tab(38); FormatCurrency(costfloss)
End If
'all if statements used only if user inputs a positive value.
'program will then print list of items purchased to user.
picOutput.Print "***************************************************"

subtotal = costbuns + costbread + costpie + costpizza + costshrimp + coststeak + costdeli + costcheese + costsalmon + costapples + costoranges + costbananas + costcarrots + costtomatoes + costlettuce + costtoothpaste + costshampoo + costfloss
picOutput.Print "Subtotal:"; Tab(38); FormatCurrency(subtotal)
'calculating and displaying the subtotal to user
End Sub


Private Sub cmdNext_Click()
frmCheckOut.Hide
frmMilage.Show
'move to Milage form
End Sub

Private Sub CmdQuit_Click()
End
End Sub


