VERSION 5.00
Begin VB.Form frmProducts 
   Caption         =   "Products"
   ClientHeight    =   10605
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   Picture         =   "frmProducts.frx":0000
   ScaleHeight     =   10605
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picData 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   9000
      ScaleHeight     =   2715
      ScaleWidth      =   3795
      TabIndex        =   19
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CommandButton cmdPartner 
      BackColor       =   &H000000FF&
      Caption         =   "Check Our Partner Store"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtPartner 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Text            =   "Don't have what you're looking for?"
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox txtMisc 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   10200
      TabIndex        =   16
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton cmdPurchase 
      Caption         =   "Purchase Miscellaneous Items"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12960
      TabIndex        =   14
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort Misc By Price"
      BeginProperty Font 
         Name            =   "Bell Gothic Std Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort Misc Alphabetically"
      BeginProperty Font 
         Name            =   "Bell Gothic Std Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdMisc 
      Caption         =   "Show Miscellaneous Items"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   2655
   End
   Begin VB.PictureBox picShow 
      Height          =   4095
      Left            =   11400
      ScaleHeight     =   4035
      ScaleWidth      =   3675
      TabIndex        =   10
      Top             =   5160
      Width           =   3735
   End
   Begin VB.TextBox txtRegister 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   9
      Text            =   "Register"
      Top             =   4200
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   360
      ScaleHeight     =   4035
      ScaleWidth      =   6315
      TabIndex        =   8
      Top             =   5160
      Width           =   6375
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   12960
      TabIndex        =   7
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmdCorkedBats 
      Caption         =   "Corked Bats"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdBalls 
      Caption         =   "Balls"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   11160
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   3600
      TabIndex        =   4
      Top             =   9600
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Go Back To Home"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   8280
      TabIndex        =   3
      Top             =   9600
      Width           =   2655
   End
   Begin VB.CommandButton cmdSteriods 
      Caption         =   "Steroids"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdGloves 
      Caption         =   "Gloves"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdCleats 
      Caption         =   "Cleats"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblPrompt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Miscellaneous Item:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   15
      Top             =   1200
      Width           =   2535
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Use the arrays, item and price.
Dim runningTotal As Single, items(1 To 7) As String, Price(1 To 7) As Single, CTR As Integer

Option Explicit

Private Sub cmdAlpha_Click()
Dim Pass As Integer, pos As Integer, Temp As Single, Temp2 As String, x As Integer
'Use a bubble sort and for/ next loops, and if then statement to sort the product by price, from least to greatest.
picData.Cls
picData.Print "items"; Tab(30); "price"; Tab(30)
picData.Print "............................................................."
For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If items(pos) > items(pos + 1) Then
            Temp = Price(pos)
            Price(pos) = Price(pos + 1)
            Price(pos + 1) = Temp
            Temp2 = items(pos)
            items(pos) = items(pos + 1)
            items(pos + 1) = Temp2
        End If
    Next pos
Next Pass
For x = 1 To CTR
    picData.Print items(x); Tab(30); FormatCurrency(Price(x))
Next x
End Sub

Private Sub cmdBalls_Click(Index As Integer)
Dim Balls As Single
'Use dynamic picture loading to load a picture of each of our products
picShow.Picture = LoadPicture(App.Path & "\Baseballs.jpg")
Balls = 4#
'Print the item and its price in a picture box.
picResults.Print "Balls", , FormatCurrency(Balls)
'Add the amount of each product to the running total.
runningTotal = runningTotal + Balls
End Sub
Private Sub cmdClear_Click(Index As Integer)
'clear the first picture box, the running total, and the second picture box.
picResults.Cls
runningTotal = 0
End Sub
Private Sub cmdCleats_Click(Index As Integer)
Dim Cleats As Single
picShow.Picture = LoadPicture(App.Path & "\cleat.jpg")
Cleats = 40#
picResults.Print "Cleats", , FormatCurrency(Cleats)
runningTotal = runningTotal + Cleats
End Sub
Private Sub cmdCorkedBats_Click(Index As Integer)
Dim CorkedBats As Single
picShow.Picture = LoadPicture(App.Path & "\bat.jpg")
CorkedBats = 275#
picResults.Print "CorkedBats", , FormatCurrency(CorkedBats)
runningTotal = runningTotal + CorkedBats
End Sub
Private Sub cmdGloves_Click(Index As Integer)
Dim Gloves As Single
picShow.Picture = LoadPicture(App.Path & "\Gloves.jpg")
Gloves = 95#
picResults.Print "Gloves", , FormatCurrency(Gloves)
runningTotal = runningTotal + Gloves
End Sub

Private Sub cmdMisc_Click()
'Use a file input to open a text file and post it in a picture box, using a do until loop.
Open App.Path & "\misc.txt" For Input As #1
picData.Print "items"; Tab(30); "price"
picData.Print ".........................................................."
CTR = 0
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, items(CTR), Price(CTR)
    picData.Print items(CTR); Tab(30); FormatCurrency(Price(CTR))
Loop



End Sub

Private Sub cmdPartner_Click()
'Use the arrays, Partner Products and Partner Price.
Dim PartnerProducts(1 To 20) As String
    Dim Search As String
    Dim CTR As Integer
    Dim Found As Boolean
    Dim PartnerPrice(1 To 100) As Single
    Dim Choice As String
'Open text file of the products carried by our partner store.
    Open App.Path & "\partnerstore.txt" For Input As #2
'Using an input box, ask the customer if there is a product that they want that our store
'doesn't carry.
    Search = InputBox("What are you looking for?", "Need Help?")
'Use the match/stop search to search our partner's product inventory to see if they carry
'the product requested by our customer.
    ' match / stop search to find out if our customer's request is in our partner's inventory.
    CTR = 0
    Found = False
    
    Do Until EOF(2) Or Found = True
        CTR = CTR + 1
        Input #2, PartnerProducts(CTR), PartnerPrice(CTR)
        If PartnerProducts(CTR) = Search Then
            Found = True
        End If
    Loop
    Close #2
 'If our partner store does carry their product, give the customer the opportunity to tell
 'us using an input  box if they want to purchase the product. If they do, print the product
 'and price in a picture box and add the price to the running total. Use a message box to
 'tell them they can pick up their product at our register. If we have product, let them
 'know.
 
    If Found = True Then
        MsgBox "Yes, they have that product!"
        Choice = InputBox("Would you like to buy this product for " & FormatCurrency(PartnerPrice(CTR)) & " , yes or no?")
        If Choice = "yes" Then
            picResults.Print PartnerProducts(CTR); Tab(30); FormatCurrency(PartnerPrice(CTR))
            runningTotal = runningTotal + PartnerPrice(CTR)
            MsgBox "Ok, you can pay for it at the register and you can pick it up at our partner store."
            End If
    Else
        MsgBox "No, they do not have that product!"
    End If


End Sub

Private Sub cmdPrice_Click()
Dim Pass As Integer, pos As Integer, Temp As String, Temp2 As Single, x As Integer
picData.Cls
picData.Print "items"; Tab(30); "price"; Tab(30)
picData.Print "............................................................."
For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If Price(pos) > Price(pos + 1) Then
            Temp = items(pos)
            items(pos) = items(pos + 1)
            items(pos + 1) = Temp
            Temp2 = Price(pos)
            Price(pos) = Price(pos + 1)
            Price(pos + 1) = Temp2
        End If
    Next pos
Next Pass
For x = 1 To CTR
    picData.Print items(x); Tab(30); FormatCurrency(Price(x))
Next x
End Sub

Private Sub cmdPurchase_Click()
'Use a text box to allow the user to chose one of our miscellaneous items, print the product
' and price in the picture box and add the price to the running total. If we do not have the
'product use a message box to tell them that we do not.
Dim miscellaneous As String
miscellaneous = txtMisc.Text
    Select Case miscellaneous
        Case Is = "Jerseys"
            picResults.Print "Jerseys"; Tab(30); FormatCurrency(60)
            runningTotal = runningTotal + 60
        Case Is = "Field of Dreams(VHS)"
            picResults.Print "Field of Dreams(VHS)"; Tab(30); FormatCurrency(10)
            runningTotal = runningTotal + 10
        Case Is = "Big League Chew"
            picResults.Print "Big League Chew"; Tab(30); FormatCurrency(3.5)
            runningTotal = runningTotal + 3.5
        Case Is = "Chewing Tobacco"
            picResults.Print "Chewing Tobacco"; Tab(30); FormatCurrency(4)
            runningTotal = runningTotal + 4
        Case Is = "Seeds"
            picResults.Print "Seeds"; Tab(30); FormatCurrency(1.25)
            runningTotal = runningTotal + 1.25
        Case Is = "Peanuts"
            picResults.Print "Peanuts"; Tab(30); FormatCurrency(2)
            runningTotal = runningTotal + 2
        Case Is = "Striped Socks"
            picResults.Print "Striped Socks"; Tab(30); FormatCurrency(6)
            runningTotal = runningTotal + 6
        Case Else
            MsgBox "We do not carry that item!"
    End Select
End Sub

Private Sub cmdQuit_Click(Index As Integer)
'Leave product page and go back to homepage.
frmHome.Show
frmProducts.Hide
End Sub
Private Sub cmdSteriods_Click(Index As Integer)
Dim Steriods As Single
picShow.Picture = LoadPicture(App.Path & "\steroid.jpg")
Steriods = 95#
picResults.Print "Steriods", , FormatCurrency(Steriods)
runningTotal = runningTotal + Steriods
End Sub
Private Sub cmdTotal_Click(Index As Integer)
Dim Tax As Single
Dim Total As Single
Dim Savings As Single
'Use if statement to determine if their trivia score, and if they do add it to the running
' total and add up their savings. Calculate the tax, print the total and the amount saved in
'picture box. Use A file output to print the total on a separate note.
If NumMatches >= 7 Then
    runningTotal = runningTotal * 0.9
    Savings = runningTotal * 0.1
End If
Tax = runningTotal * 0.15
Total = runningTotal + Tax
picResults.Print "----------------------------"
picResults.Print "Subtotal", , FormatCurrency(runningTotal)
picResults.Print "Tax", , FormatCurrency(Tax)
picResults.Print "Total", , FormatCurrency(Total)
picResults.Print "You saved ", , FormatCurrency(Savings)
Open App.Path & "\Output.txt" For Output As #3
    Write #3, "Subtotal:" & FormatCurrency(runningTotal)
    Write #3, "Tax:" & FormatCurrency(Tax)
    Write #3, "Total:" & FormatCurrency(Total)
    Write #3, "Total:" & FormatCurrency(Total)
Close #3
    
End Sub




