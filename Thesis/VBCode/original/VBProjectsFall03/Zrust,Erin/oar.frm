VERSION 5.00
Begin VB.Form frmOARCDs 
   BackColor       =   &H000040C0&
   Caption         =   "O.A.R. CDs"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optWanderer 
      Caption         =   "Option4"
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optRisen 
      Caption         =   "Option3"
      Height          =   255
      Left            =   6240
      TabIndex        =   14
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optAnyTime 
      Caption         =   "Option2"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton optInBetween 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0000C000&
      Caption         =   "Next band/songwriter"
      Height          =   735
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000C000&
      Caption         =   "Add to your shopping cart"
      Height          =   735
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   2895
   End
   Begin VB.PictureBox Out 
      BackColor       =   &H0000C000&
      Height          =   3495
      Left            =   600
      ScaleHeight     =   3435
      ScaleWidth      =   6555
      TabIndex        =   9
      Top             =   3840
      Width           =   6615
   End
   Begin VB.PictureBox imgWanderer 
      Height          =   1935
      Left            =   7920
      Picture         =   "oar.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox imgRisen 
      Height          =   1935
      Left            =   5400
      Picture         =   "oar.frx":65DC
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox imgAnyTime 
      Height          =   1935
      Left            =   2760
      Picture         =   "oar.frx":B0A1
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox imgInBetween 
      Height          =   1935
      Left            =   360
      Picture         =   "oar.frx":EF17
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblSelectOARCD 
      BackColor       =   &H000040C0&
      Caption         =   "Select an O.A.R. CD"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblWanderer 
      BackColor       =   &H000040C0&
      Caption         =   "The Wanderer"
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblRisen 
      BackColor       =   &H000040C0&
      Caption         =   "Risen"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblAnyTime 
      BackColor       =   &H000040C0&
      Caption         =   "Any Time Now"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblInBetween 
      BackColor       =   &H000040C0&
      Caption         =   "In Between Now and Then"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "frmOARCDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjAmazonCDPurchase (Erin Zrust's VB Project.vpb)
'Form Name : frmOARCDs (oar.frm)
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
If optInBetween = True Then
    A = 1
ElseIf optAnyTime = True Then
    A = 2
ElseIf optRisen = True Then
    A = 3
ElseIf optWanderer = True Then
    A = 4
End If
'Print the name of CD, price and ranking from data file
Out.Print DaveMatthewsBandName(D); Tab(35); FormatCurrency(DaveMatthewsBandPrice(D)); Tab(55); DaveMatthewsBandRanking(D)
Out.Print BenHarperName(B); Tab(35); FormatCurrency(BenHarperPrice(B)); Tab(55); BenHarperRanking(B)
Out.Print OARName(A); Tab(35); FormatCurrency(OARPrice(A)); Tab(55); OARRanking(A)
'Hide the "Add to your shopping cart" button
'and show the "Next Singer/Songwriter" button.
cmdAdd.Enabled = False
cmdNext.Enabled = True
End Sub

Private Sub cmdNext_Click()
'Hide the O.A.R. CD selection screen and show
'the Jack Johnson selection screen for the user's next selection.
frmOARCDs.Hide
frmJackJohnsonCDs.Show

End Sub


Private Sub Form_Load()
'ReDim Arrays that were made Public in ModuleAmazonCDSelection
'(ModAmazonCDSelection.bas) and define them for frmOARCDs.
ReDim OARName(1 To 4) As String
ReDim OARPrice(1 To 4) As Double
ReDim OARRanking(1 To 4) As Integer

'Open the data file "oar.txt" for the Arrays that
'are used in frmOARCDs.
Open PATH & "oar.txt" For Input As #1
    For A = 1 To 4
        Input #1, OARName(A), OARPrice(A), OARRanking(A)
    Next A
Close #1
'Hide "Add to your shopping cart" and
'"Next Singer/Songwriter" buttons
cmdAdd.Enabled = False
cmdNext.Enabled = False
End Sub

Private Sub optInBetween_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optAnyTime_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optRisen_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

Private Sub optWanderer_Click()
'Show "Add to shopping cart" button after a selection has been made.
cmdAdd.Enabled = True
End Sub

