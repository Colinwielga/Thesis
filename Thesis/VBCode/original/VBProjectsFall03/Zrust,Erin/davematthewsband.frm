VERSION 5.00
Begin VB.Form frmDaveMatthewsBandCDs 
   BackColor       =   &H000040C0&
   Caption         =   "Dave Matthews Band CDs"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optBS 
      Caption         =   "Option4"
      Height          =   255
      Left            =   7920
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optLiC 
      Caption         =   "Option3"
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optBTCS 
      Caption         =   "Option2"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optCrash 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0000C000&
      Caption         =   "Next Band/Songwriter"
      Height          =   735
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000C000&
      Caption         =   "Add to Your Shopping Cart"
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   2295
   End
   Begin VB.PictureBox Out 
      BackColor       =   &H0000C000&
      Height          =   4215
      Left            =   600
      ScaleHeight     =   4155
      ScaleWidth      =   7635
      TabIndex        =   4
      Top             =   3720
      Width           =   7695
   End
   Begin VB.PictureBox imgBS 
      Height          =   1815
      Left            =   7200
      Picture         =   "davematthewsband.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.PictureBox imgLiC 
      Height          =   1815
      Left            =   4800
      Picture         =   "davematthewsband.frx":6FE3
      ScaleHeight     =   1755
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.PictureBox imgBTCS 
      Height          =   1935
      Left            =   2520
      Picture         =   "davematthewsband.frx":CA17
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox imgCrash 
      Height          =   1935
      Left            =   360
      Picture         =   "davematthewsband.frx":130FE
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblBS 
      BackColor       =   &H000040C0&
      Caption         =   "Busted Stuff"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblLiC 
      BackColor       =   &H000040C0&
      Caption         =   "Live in Chicago- United Center"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblBTCS 
      BackColor       =   &H000040C0&
      Caption         =   "Before These Crowded Streets"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblCrash 
      BackColor       =   &H000040C0&
      Caption         =   "Crash"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblSelectDMB 
      BackColor       =   &H000040C0&
      Caption         =   "Select a Dave Matthews Band CD"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
End
Attribute VB_Name = "frmDaveMatthewsBandCDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjAmazonCDPurchase (Erin Zrust's VB Project.vpb)
'Form Name : frmDaveMatthewsBandCDs (davematthewsband.frm)
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
If optCrash = True Then
    D = 1
ElseIf optBTCS = True Then
    D = 2
ElseIf optLiC = True Then
    D = 3
ElseIf optBS = True Then
    D = 4
End If
'Print the name of CD, price and ranking from data file
Out.Print DaveMatthewsBandName(D);
Out.Print Tab(35); FormatCurrency(DaveMatthewsBandPrice(D));
Out.Print Tab(55); DaveMatthewsBandRanking(D)

'Hide the "Add to your shopping cart" button
'and show the "Next Singer/Songwriter" button.
cmdAdd.Enabled = False
cmdNext.Enabled = True

End Sub

Private Sub cmdNext_Click()
'Hide the Dave Matthews Band CD selection screen and show
'the Ben Harper selection screen for next selection.
frmDaveMatthewsBandCDs.Hide
frmBenHarperCDs.Show

End Sub


Private Sub Form_Load()
PATH = "M:\CS130\Project1\"
'ReDim Arrays that were made Public in ModuleAmazonCDSelection
'(ModAmazonCDSelection.bas) and define them for frmBenHarperCDs.
ReDim DaveMatthewsBandName(1 To 4) As String
ReDim DaveMatthewsBandPrice(1 To 4) As Double
ReDim DaveMatthewsBandRanking(1 To 4) As Integer

'Open the data file "davematthewsband.txt" for the Arrays that
'are used in frmDaveMatthewsBandCDs.
Open PATH & "davematthewsband.txt" For Input As #1
    For D = 1 To 4
        Input #1, DaveMatthewsBandName(D), DaveMatthewsBandPrice(D), DaveMatthewsBandRanking(D)
    Next D
Close #1

'Hide "Add to your shopping cart" and
'"Next Singer/Songwriter" buttons
'this form cannot be skipped when the
'program begins
cmdAdd.Enabled = False
cmdNext.Enabled = False
End Sub

Private Sub optBS_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optBTCS_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optCrash_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optLiC_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub
