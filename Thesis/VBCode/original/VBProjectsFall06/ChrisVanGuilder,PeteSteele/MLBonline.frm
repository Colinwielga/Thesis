VERSION 5.00
Begin VB.Form frmGear 
   Caption         =   "Baseball Gear"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   7560
      Width           =   4455
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Homepage"
      Height          =   2055
      Left            =   7080
      TabIndex        =   2
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdCart 
      Caption         =   "Send my Baseball gear purchases to My Cart"
      Height          =   2055
      Left            =   4680
      TabIndex        =   1
      Top             =   5040
      Width           =   2295
   End
   Begin VB.PictureBox picResultsGear 
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2715
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Image imgSpikes 
      Height          =   1665
      Left            =   6840
      Picture         =   "MLBonline.frx":0000
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1800
   End
   Begin VB.Image imgShorts 
      Height          =   1515
      Left            =   6720
      Picture         =   "MLBonline.frx":3C9C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1830
   End
   Begin VB.Image imgHelmet 
      Height          =   1440
      Left            =   3480
      Picture         =   "MLBonline.frx":59A6
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Image imgBag 
      Height          =   1185
      Left            =   120
      Picture         =   "MLBonline.frx":7BF7
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   2460
   End
   Begin VB.Image imgBall 
      Height          =   1515
      Left            =   3480
      Picture         =   "MLBonline.frx":E680
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1635
   End
   Begin VB.Image imgGlove 
      Height          =   1170
      Left            =   120
      Picture         =   "MLBonline.frx":11010
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1650
   End
   Begin VB.Image imgBat 
      Height          =   1500
      Left            =   120
      Picture         =   "MLBonline.frx":14C73
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "frmGear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCart_Click()
    frmCart.Visible = True                 'sends accessory total to cart
    frmGear.Visible = False
    frmCart.picResults.Print "Gear Purchases", "Total"
    frmCart.picResults.Print "*************************************************"
    frmCart.picResults.Print "You Gear Total is: ", FormatCurrency(GearSum)
End Sub
Private Sub cmdClear_Click() 'clears accessory total
picResultsGear.Cls
GearSum = 0
End Sub
Private Sub cmdReturn_Click() 'returns to HomePage
frmGear.Visible = False
frmHomepage.Visible = True
End Sub
Private Sub cmdTotalA_Click()  'gets total for accessories and prints in accessories form
picResultsGear.Print "******************************************************"
picResultGear.Print "Total Gear: ", FormatCurrency(Accsum, 2)
cmdCart.Visible = True
End Sub

Private Sub imgBat_Click()    'buys Gear and adds cost to total
picResultsGear.Print "Louisville Slugger Bat(s)   ", FormatCurrency(99, 2)
GearSum = GearSum + 99
End Sub

Private Sub imgGlove_Click() 'buys gear and adds cost to total
picResultsGear.Print "Official Baseball Glove", FormatCurrency(79.99, 2)
GearSum = GearSum + 79.99
End Sub

Private Sub imgBag_Click() 'buys gear and adds cost to total
picResultsGear.Print "Baseball Gear Bag", FormatCurrency(49.99, 2)
GearSum = GearSum + 49.99
End Sub

Private Sub imgHelmet_Click()   'buys gear and adds cost to total
picResultsGear.Print "Tama Drum Heads", FormatCurrency(45.99, 2)
GearSum = GearSum + 45.99
End Sub

Private Sub imgBall_Click()
Dim Number As Integer
Number = InputBox("enter the number of baseballs you would like to order in multiples of 10 between 10 and 30. Purchase 30 to receive $10 discount", "Place Order")
Select Case Number                'input number of baseballs you would like
Case Is = 10
    picResultsGear.Print "Official MLB Baseball", FormatCurrency(20, 2)
    GearSum = GearSum + 10          'buys 10 baseballs and adds cost to total
Case 20
    picResultsGear.Print "Official MLB Baseball", FormatCurrency(40, 2)
    GearSum = GearSum + 20          'buys 20 baseballs and adds cost to total
Case 30
    picResultsGear.Print "Official MLB Baseball", FormatCurrency(50, 2)
    GearSum = GearSum + 30          'buys 30 baseballs and adds cost to total
Case Else
    MsgBox "Please enter a multiple of 10 between 10 and 30", , "Invalid Request"
End Select                        'if a wrong quantity is typed display invalid message
End Sub

Private Sub imgShorts_Click()   'buys gear and adds cost to total
picResultsGear.Print "UnderArmour Sliding Pants", FormatCurrency(49.99, 2)
GearSum = GearSum + 49.99
End Sub
Private Sub imgSpikes_Click()
    Dim InchesArray(1 To 25) As Integer
    Dim NamesArray(1 To 12) As String
    Dim PricesArray(1 To 300) As Single
    Dim Found As Boolean
    Dim Pos, Search As Integer
        
    Open App.Path & "\spikes.txt" For Input As #1 'opens spikes size file and places into three arrays
        Pos = 0
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, InchesArray(Pos), NamesArray(Pos), PricesArray(Pos)
    Loop
    Close #1
    
    Found = False
    Pos = 0
    Search = InputBox("enter the size (9-12 mens foot size) of baseball spikes you would like", "Spikes Size")
        
    Do While Found = False
        Pos = Pos + 1
        If Search = InchesArray(Pos) Then
        Found = True
        End If
    Loop
    If Found = True Then
        picResultsGear.Print NamesArray(Pos), FormatCurrency(PricesArray(Pos))
        GearSum = GearSum + PricesArray(Pos)
    End If
End Sub
