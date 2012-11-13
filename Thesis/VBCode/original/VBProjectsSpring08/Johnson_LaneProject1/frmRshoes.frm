VERSION 5.00
Begin VB.Form frmRshoes 
   Caption         =   "Running Shoes"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   Picture         =   "frmRshoes.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   375
      Left            =   4080
      TabIndex        =   26
      Top             =   6000
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   1575
      Left            =   3000
      ScaleHeight     =   1515
      ScaleWidth      =   6675
      TabIndex        =   25
      Top             =   4080
      Width           =   6735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   9000
      TabIndex        =   16
      Top             =   8280
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   255
      Left            =   6720
      TabIndex        =   15
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   255
      Left            =   10800
      TabIndex        =   12
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   255
      Left            =   10680
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   255
      Left            =   10560
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   5040
      Picture         =   "frmRshoes.frx":2CC80
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   7200
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   2880
      Picture         =   "frmRshoes.frx":2D5D0
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   7200
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   600
      Picture         =   "frmRshoes.frx":2DF1C
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   10320
      Picture         =   "frmRshoes.frx":2EABC
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   600
      Picture         =   "frmRshoes.frx":2F484
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   10320
      Picture         =   "frmRshoes.frx":2FF98
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   10320
      Picture         =   "frmRshoes.frx":30A62
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   600
      Picture         =   "frmRshoes.frx":3149E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   9960
      TabIndex        =   20
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   9960
      TabIndex        =   19
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   9960
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "7"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   7320
      Width           =   255
   End
End
Attribute VB_Name = "frmRshoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmRshoes
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the men's running shoes form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Option Explicit
Dim Rshoe(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmRshoes.Hide
frmShoes.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\RshoeArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Rshoe(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Rshoe(j), Tab(25); Prices(j)
    End If
Next j
    If n > j Then 'this loops gives an error message of the user enters a number that doesnt correspond with the labeled items on the form
        MsgBox "Oooops! You have Entered an invalid Number. Please enter a valid number"
    End If
    
'this loop will keep the running total of items and make it viewable to the usersFor i = 1 To ctr
 For i = 1 To ctr
    If n = i Then
        found = True
        sum = sum + Prices(i)
        picResults.Print Tab(25); Tab(50); sum  'prints the users running total
    End If
 Next i

End Sub

Private Sub cmdInfo_Click() 'allows the user to view the specific information on the item
MsgBox "The Nike Air Max 360 II SL features the most cushioning ever engineered into a running shoe for the most comfortable ride ever. Breathable synthetic leather upper with supportive rand and 360 degree reflectivity. Unique full-length Air-Sole® provides durability, comfort and unsurpassed ride. Durable BRS 1000® rubber outsole with deep forefoot flex grooves offers enhanced flexibility.", , "Nike Men's Air Max 360 II SL"
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Max TL IV running shoe is made for the runner seeking ride, durability and comfort. Lightweight synthetic mesh upper offers comfort, fit and breathability. A full-length Air-Sole® cushioning unit provides maximum impact protection. The BRS 1000™ carbon rubber outsole adds great traction in all conditions. Wt. 15.4 oz.", , "Nike Men's Air Max TL IV Premier"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Max TL IV running shoe is made for the runner seeking ride, durability and comfort. Lightweight synthetic mesh upper offers comfort, fit and breathability. A full-length Air-Sole® cushioning unit provides maximum impact protection. The BRS 1000™ carbon rubber outsole adds great traction in all conditions. Wt. 15.4 oz.", , "Nike Men's Air Max TL IV SL"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Shox Turbo+ VI running shoe is made for the runner who needs the best cushioning that footwear can provide. Modern, flexible midfoot strapping system creates a snug fit around the midfoot. Light and breathable mesh upper offers a great performance feel. Nike Shox™ cushioning system in the heel. Flexible and well-cushioned Phylon™ forefoot provides a nice responsive ride. Full-length BRS 1000™ carbon rubber outsole adds great durability. Wt. 13 oz.", , "Nike Men's Shox Turbo + VI"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Max TL 3 running shoe is made for the runner seeking ride, durability and comfort. Lightweight synthetic mesh upper offers comfort, fit and breathability. A full-length Air-Sole® cushioning unit provides maximum impact protection. The BRS 1000™ carbon rubber outsole adds great traction in all conditions.", , "Nike Men's Air Max TL 3"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Max International is built for the runner seeking a cushioned, supportive running shoe. Full-grain leather upper. Large-volume heel Air-Sole® unit with polyurethane. Solid rubber outsole. Wt. 14.6 oz.", , "Nike Men's Air Max International"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "The coolest running shoe to hit the streets, the Nike Zoom Jasari+ features minimal mesh panels for maximum performance. Full-length Phylon™ midsole with Zoom Air™ unit in the heel and forefoot adds responsive feel. Lightweight mesh upper provides superior breathability. Full-length rubber outsole with Nike Waffle® design. Nike+ ready. Wt. 8.4 oz.", , "Nike Men's Zoom Jasari +"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Free 7.0 is a highly responsive running shoe that helps build foot and leg strength. Synthetic mesh upper with overlays gives a comfortable and supportive fit. CEVA midsole with Free sipes offers advanced foot strengthening. Strategically placed BRS 1000™ provides enhanced durability. Wt. 9.2 oz.", , "Nike Men's Free 7.0"
End Sub



