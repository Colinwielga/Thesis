VERSION 5.00
Begin VB.Form frmPrices 
   BackColor       =   &H00C000C0&
   Caption         =   "Prices of Flowers"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   FillColor       =   &H0000FFFF&
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFlower4 
      Height          =   1335
      Left            =   8640
      Picture         =   "frmPrices.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   6840
      Width           =   1935
   End
   Begin VB.PictureBox picFlower3 
      Height          =   1335
      Left            =   6720
      Picture         =   "frmPrices.frx":83F2
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   6840
      Width           =   1935
   End
   Begin VB.PictureBox PicFlower2 
      Height          =   1335
      Left            =   4800
      Picture         =   "frmPrices.frx":107E4
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   6840
      Width           =   1935
   End
   Begin VB.PictureBox picFlower1 
      Height          =   1335
      Left            =   2880
      Picture         =   "frmPrices.frx":18BD6
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H008080FF&
      Caption         =   "Search for Prices"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H008080FF&
      Caption         =   "Go Back to Main Menu"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H008080FF&
      Caption         =   "Display Prices"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H008080FF&
      Caption         =   "Load Prices"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdSortPrices 
      BackColor       =   &H008080FF&
      Caption         =   "Flowers: Most Expensive to Least Expensive"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   4440
      ScaleHeight     =   4875
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Designed By Allison Becker"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   8040
      Width           =   3255
   End
   Begin VB.Label lblPrices 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Prices of Flowers"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1095
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "frmPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Flowers For U! (FlowerShop.vbp)
'Form Name: (frmPrices)
'Author: Allison Becker
'Date Written: 3/23/06
'Objective: The objective of this page is to show the user the prices that we
'charge for each flower we offer. You are able to search for prices that
'you are willing to pay. You are also able to look at the products in a
'price list randomly and from the least expensive to the most expensive,
'to make it easier to decide what you may want to purchase.

Option Explicit
Dim Pos As Integer

Private Sub cmdDisplay_Click()
    picResults.Cls
    'displays results to picture box after either the bubble sort or exhaustive search
    picResults.Print "Flowers"; Tab(25); "Price"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Flowers(Pos); Tab(25); FormatCurrency(Prices(Pos))
    Next Pos
End Sub

Private Sub cmdLoad_Click()
  Pos = 0
    Open App.Path & "\Flowers.txt" For Input As #1 'Opens txt file and reads the information
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Flowers(Pos), Prices(Pos)
    Loop
    Close #1
    Size = Pos
End Sub

Private Sub cmdMain_Click()
    frmPrices.Hide
    frmFlowerShop.Show
End Sub

Private Sub cmdSearch_Click()
    Dim Found As Boolean
    Dim Pos As Integer
    Dim SearchValue As String
    'Exhaustive Search, searching through the array to find a price in accordance to the price entered
    SearchValue = InputBox("Indicate Price Willing to Pay", "Input") 'taking input from an input box
    Found = False
    picResults.Cls
    For Pos = 1 To Size
        If Prices(Pos) = SearchValue Then 'printing results to a picture box
            picResults.Print "Flowers within price rage are " & Flowers(Pos)
            Found = True
        End If
    Next Pos
    If Found = False Then
        MsgBox "No Matches Found!", vbCritical, "Failure"
    End If
    
End Sub

Private Sub cmdSortPrices_Click()
    Dim TempFlowers, TempPrices As Single
    Dim Pass, Pos As Integer
    'Bubble Sort of the prices
    For Pass = 1 To (Size - 1)
        For Pos = 1 To (Size - Pass)
            If Prices(Pos) > Prices(Pos + 1) Then
                TempFlowers = Flowers(Pos)
                Flowers(Pos) = Flowers(Pos + 1)
                Flowers(Pos + 1) = TempFlowers
                TempPrices = Prices(Pos)
                Prices(Pos) = Prices(Pos + 1)
                Prices(Pos + 1) = TempPrices
            End If
        Next Pos
    Next Pass

End Sub

