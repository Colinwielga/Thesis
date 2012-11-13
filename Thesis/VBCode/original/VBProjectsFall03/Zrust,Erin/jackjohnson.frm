VERSION 5.00
Begin VB.Form frmJackJohnsonCDs 
   BackColor       =   &H000040C0&
   Caption         =   "Jack Johnson CDs"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optBonnaroo 
      Caption         =   "Option4"
      Height          =   255
      Left            =   9000
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optSeptember 
      Caption         =   "Option3"
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optOnandOn 
      Caption         =   "Option2"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton optBrushfire 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0000C000&
      Caption         =   "Proceed to checkout"
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000C000&
      Caption         =   "Add to your shopping cart"
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   2775
   End
   Begin VB.PictureBox out 
      BackColor       =   &H0000C000&
      Height          =   3975
      Left            =   480
      ScaleHeight     =   3915
      ScaleWidth      =   6915
      TabIndex        =   9
      Top             =   3360
      Width           =   6975
   End
   Begin VB.PictureBox imgBonnaroo 
      Height          =   1935
      Left            =   8160
      Picture         =   "jackjohnson.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox imgSeptember 
      Height          =   1815
      Left            =   5640
      Picture         =   "jackjohnson.frx":36E5
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox imgOnandOn 
      Height          =   1815
      Left            =   2880
      Picture         =   "jackjohnson.frx":8682
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox imgBrushfire 
      Height          =   1815
      Left            =   360
      Picture         =   "jackjohnson.frx":CD78
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblSelectJackCD 
      BackColor       =   &H000040C0&
      Caption         =   "Select a Jack Johnson CD"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblBonnaroo 
      BackColor       =   &H000040C0&
      Caption         =   "Live from Bonnaroo Music Festival "
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblSeptember 
      BackColor       =   &H000040C0&
      Caption         =   "The September Sessions"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblOnandOn 
      BackColor       =   &H000040C0&
      Caption         =   "On and On"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblBrushfire 
      BackColor       =   &H000040C0&
      Caption         =   "Brushfire Fairytales"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "frmJackJohnsonCDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjAmazonCDPurchase (Erin Zrust's VB Project.vpb)
'Form Name : frmJackJohnsonCDs (jackjohnson.frm)
'Author: Erin Zrust
'Date Written: October 30, 2003
'Purpose of Form: To purchase various CDS, find the total
                  'including shipping and tax, the Amazon.com
                  'ranking for each CD purchase and the average
                  'of all CDs purchased.

'This forces user to declare all variables
Option Explicit


Private Sub cmdAdd_Click()
'Print out headings
Out.Print "Name of CD"; Tab(35); "Price"; Tab(55); "Amazon.com ranking"
Out.Print "----------------------------------------------------------------------------------------------------------------------"
'Find out what CD the user chose
If optBrushfire = True Then
    J = 1
ElseIf optOnandOn = True Then
    J = 2
ElseIf optSeptember = True Then
    J = 3
ElseIf optBonnaroo = True Then
    J = 4
End If
'Print the name of CD, price and ranking from data file
Out.Print DaveMatthewsBandName(D); Tab(35); FormatCurrency(DaveMatthewsBandPrice(D)); Tab(55); DaveMatthewsBandRanking(D)
Out.Print BenHarperName(B); Tab(35); FormatCurrency(BenHarperPrice(B)); Tab(55); BenHarperRanking(B)
Out.Print OARName(A); Tab(35); FormatCurrency(OARPrice(A)); Tab(55); OARRanking(A)
Out.Print JackJohnsonName(J); Tab(35); FormatCurrency(JackJohnsonPrice(J)); Tab(55); JackJohnsonRanking(J)
'Hide the "Add to your shopping cart" button
'and show the "Proceed to checkout" button.
cmdAdd.Enabled = False
cmdNext.Enabled = True


End Sub

Private Sub cmdNext_Click()
'Hide the Jack Johnson CD selection screen and show
'the Finish Checkout selection screen for the user's next selection.
frmJackJohnsonCDs.Hide
frmFinishCheckout.Show

End Sub


Private Sub Form_Load()
'ReDim Arrays that were made Public in ModuleAmazonCDSelection
'(ModAmazonCDSelection.bas) and define them for frmJackJohnsonCDs.
ReDim JackJohnsonName(1 To 4) As String
ReDim JackJohnsonPrice(1 To 4) As Double
ReDim JackJohnsonRanking(1 To 4) As Integer

'Open the data file "jackjohnson.txt" for the Arrays that
'are used in frmJackJohnsonCDs.
Open PATH & "jackjohnson.txt" For Input As #1
    For J = 1 To 4
        Input #1, JackJohnsonName(J), JackJohnsonPrice(J), JackJohnsonRanking(J)
    Next J
Close #1

'Hide "Add to your shopping cart" and
'"Next Singer/Songwriter" buttons
cmdAdd.Enabled = False
cmdNext.Enabled = False
End Sub

Private Sub optBrushfire_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optOnandOn_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optSeptember_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optBonnaroo_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub


