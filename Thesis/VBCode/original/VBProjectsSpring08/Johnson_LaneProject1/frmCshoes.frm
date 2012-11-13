VERSION 5.00
Begin VB.Form frmCshoes 
   Caption         =   "Casual Shoes"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   Picture         =   "frmCshoes.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   615
      Left            =   840
      Picture         =   "frmCshoes.frx":16A84E
      TabIndex        =   26
      Top             =   6240
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      Height          =   1935
      Left            =   360
      ScaleHeight     =   1875
      ScaleWidth      =   7515
      TabIndex        =   25
      Top             =   4200
      Width           =   7575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   5280
      Picture         =   "frmCshoes.frx":16F260
      TabIndex        =   16
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   255
      Left            =   11280
      TabIndex        =   15
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   255
      Left            =   11160
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   255
      Left            =   11280
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   255
      Left            =   9000
      TabIndex        =   12
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   255
      Left            =   11400
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   255
      Left            =   9000
      TabIndex        =   10
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   255
      Left            =   8880
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   255
      Left            =   8760
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   11040
      Picture         =   "frmCshoes.frx":173C72
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   6480
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   11040
      Picture         =   "frmCshoes.frx":174581
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   10920
      Picture         =   "frmCshoes.frx":174F1B
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   10920
      Picture         =   "frmCshoes.frx":175688
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   8760
      Picture         =   "frmCshoes.frx":175E70
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   6480
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   8640
      Picture         =   "frmCshoes.frx":1768C5
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   8640
      Picture         =   "frmCshoes.frx":177245
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   8640
      Picture         =   "frmCshoes.frx":177BA9
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   10800
      TabIndex        =   24
      Top             =   6600
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   10800
      TabIndex        =   23
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   10680
      TabIndex        =   22
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   10680
      TabIndex        =   21
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   8520
      TabIndex        =   20
      Top             =   6600
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   8400
      TabIndex        =   19
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   18
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   17
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "frmCshoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmCshoe
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the men's Casual shoes form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Option Explicit
Dim Cshoes(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmCshoes.Hide
frmShoes.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()


found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\CshoeArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Cshoes(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Cshoes(j), Tab(25); Prices(j)
    End If
Next j
    If n > j Then 'this loops gives an error message of the user enters a number that doesnt correspond with the labeled items on the form
        MsgBox "Oooops! You have Entered an invalid Number. Please enter a valid number"
    End If

'this loop will keep the running total of items and make it viewable to the users
For i = 1 To ctr
    If n = i Then
        found = True
        sum = sum + Prices(i)
        picResults.Print Tab(25); Tab(50); sum  'prints the users running total
    End If
         
Next i
End Sub

Private Sub cmdInfo_Click()     'allows the user to view the specific information on the item
MsgBox "Forget sweet talking ... try sweet walking. Enjoy the divine stylings of the Nike Air Max Light featuring a feather-light leather and mesh upper, visible air unit in the heel and a sturdy rubber outsole.", , "Nike Men's Air Max Light"
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "A modern take on the classic boat shoe. The Nike Mad Jibe is a men's casual shoe with an athletically designed leather upper and a durable rubber outsole.", , "Nike Men's Mad Jibe"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "Respect your roots with the retro styling of the Nike Vandal Low Premium, a classic hoops shoe. An EVA midsole offers stable, durable cushioning, while a rubber cupsole cradles your foot and cushions the ride. The upper is leather, and a solid rubber outsole ensures great traction on the street or in the paint.", , "Nike Men's Vandal Low Premium "
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Sellwood is a crisp, clean-looking casual shoe inspired by the court. It features a leather and mesh upper with a visible Air® unit in the heel and a durable rubber outsole.", , "Nike Men's Air Sellwood"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The casual side of sport comes into play in the Nike Capri Slip. This casual, slip-on shoe features a fabric upper with a durable rubber outsole.", , "Nike Men's Capri Slip"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Blazer Low casual shoe is the original hoops shoe with a splash of pretty. Rich leather upper ensures a comfortable, smooth fit and feel. New patterns and fresh colors add hip style. Rubber outsole offers great traction.", , "Nike Men's Blazer Low"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "The next best thing to being barefoot. The Nike Men's U-Soc is a casual sandal that takes you back to the original concept ... minimal footwear for having fun in the sun and water. A quick-drying mesh upper features welded toe and heel protection. Nike 0.44 ultra-sticky rubber outsole provides Spiderman-like grip.", , "Nike Men's U-Soc"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "A richly detailed leather boot. Rich, waterproof full-grain leather upper with metal hardware. Full-length visible Air-Sole unit for the ultimate in cushioning. Solid rubber compound and lug pattern outsole for maximum traction and durability.", , "Nike Men's Air Max Goadome"
End Sub

