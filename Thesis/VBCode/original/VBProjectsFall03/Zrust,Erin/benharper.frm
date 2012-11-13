VERSION 5.00
Begin VB.Form frmBenHarperCDs 
   BackColor       =   &H000040C0&
   Caption         =   "Ben Harper CDs"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optFight 
      Caption         =   "Option4"
      Height          =   255
      Left            =   9000
      TabIndex        =   15
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optBurn 
      Caption         =   "Option3"
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optMars 
      Caption         =   "Option2"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optDiamonds 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0000C000&
      Caption         =   "Next band/songwriter"
      Height          =   735
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000C000&
      Caption         =   "Add to Shopping Cart"
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   2655
   End
   Begin VB.PictureBox Out 
      BackColor       =   &H0000C000&
      Height          =   4095
      Left            =   600
      ScaleHeight     =   4035
      ScaleWidth      =   6555
      TabIndex        =   9
      Top             =   3720
      Width           =   6615
   End
   Begin VB.PictureBox imgFightforYourMind 
      Height          =   1935
      Left            =   8040
      Picture         =   "benharper.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.PictureBox imgBurntoShine 
      Height          =   1815
      Left            =   5520
      Picture         =   "benharper.frx":4DC1
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.PictureBox imgLivefromMars 
      Height          =   1935
      Left            =   2880
      Picture         =   "benharper.frx":A15B
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.PictureBox imgDiamondsontheInside 
      Height          =   1935
      Left            =   360
      Picture         =   "benharper.frx":F18B
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblSelectBenHarper 
      BackColor       =   &H000040C0&
      Caption         =   "Select a Ben Harper CD"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblFight 
      BackColor       =   &H000040C0&
      Caption         =   "Fight for Your Mind"
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblBurn 
      BackColor       =   &H000040C0&
      Caption         =   "Burn to Shine"
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblMars 
      BackColor       =   &H000040C0&
      Caption         =   "Live From Mars"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblDiamonds 
      BackColor       =   &H000040C0&
      Caption         =   "Diamonds on the Inside"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "frmBenHarperCDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjAmazonCDPurchase (Erin Zrust's VB Project.vpb)
'Form Name : frmBenHarperCDs (benharper.frm)
'Author: Erin Zrust
'Date Written: October 30, 2003
'Purpose of Form: To purchase various CDS, find the total
                  'including shipping and tax, the Amazon.com
                  'ranking for each CD purchase and the average
                  'of all CDs purchased.

'This forces the user to declare all variables
Option Explicit



Private Sub cmdAdd_Click()
'Print out headings
Out.Print "Name of CD"; Tab(35); "Price"; Tab(55); "Amazon.com ranking"
Out.Print "----------------------------------------------------------------------------------------------------------------------"
'Find out what CD the user chose
If optDiamonds = True Then
    B = 1
ElseIf optMars = True Then
    B = 2
ElseIf optBurn = True Then
    B = 3
ElseIf optFight = True Then
    B = 4
End If
'Print the name of CD, price and ranking from data file
Out.Print DaveMatthewsBandName(D); Tab(35); FormatCurrency(DaveMatthewsBandPrice(D)); Tab(55); DaveMatthewsBandRanking(D)
Out.Print BenHarperName(B); Tab(35); FormatCurrency(BenHarperPrice(B)); Tab(55); BenHarperRanking(B)

'Hide the "Add to your shopping cart" button
'and Show the "Next Singer/Songwriter" button.
cmdAdd.Enabled = False
cmdNext.Enabled = True
End Sub

Private Sub cmdNext_Click()
'Hide the Ben Harper selection screen and show
'the O.A.R. selection screen for the user's next selection.
frmBenHarperCDs.Hide
frmOARCDs.Show

End Sub


Private Sub Form_Load()
'ReDim Arrays that were made Public in ModuleAmazonCDSelection
'(ModAmazonCDSelection.bas) and define them for frmBenHarperCDs.
ReDim BenHarperName(1 To 4) As String
ReDim BenHarperPrice(1 To 4) As Double
ReDim BenHarperRanking(1 To 4) As Integer

'Open the data file "benharper.txt" for the Arrays that
'are used in frmBenHarperCDs.
Open PATH & "benharper.txt" For Input As #1
    For B = 1 To 4
        Input #1, BenHarperName(B), BenHarperPrice(B), BenHarperRanking(B)
    Next B
Close #1

'Hide "Add to your shopping cart" and
'"Next Singer/Songwriter" buttons
cmdAdd.Enabled = False
cmdNext.Enabled = False
End Sub

Private Sub optDiamonds_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optMars_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optBurn_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optFight_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

