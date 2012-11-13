VERSION 5.00
Begin VB.Form frmWomensCasual 
   Caption         =   "Women's Casual Shoes"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   Picture         =   "frmWomensCasual.frx":0000
   ScaleHeight     =   8655
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   855
      Left            =   9240
      TabIndex        =   32
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   4680
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   2415
      Left            =   2760
      ScaleHeight     =   2355
      ScaleWidth      =   5355
      TabIndex        =   20
      Top             =   2160
      Width           =   5415
   End
   Begin VB.CommandButton cmdInfo10 
      Caption         =   "Info"
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo9 
      Caption         =   "Info"
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   255
      Left            =   9000
      TabIndex        =   16
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   255
      Left            =   9120
      TabIndex        =   14
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo1 
      Caption         =   "Info"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox Picture10 
      Height          =   1575
      Left            =   8760
      Picture         =   "frmWomensCasual.frx":18004E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   6240
      Picture         =   "frmWomensCasual.frx":1808C9
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   8760
      Picture         =   "frmWomensCasual.frx":1810FF
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   2400
      Picture         =   "frmWomensCasual.frx":181A3B
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   840
      Picture         =   "frmWomensCasual.frx":18203D
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   4320
      Picture         =   "frmWomensCasual.frx":182ABE
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   6240
      Picture         =   "frmWomensCasual.frx":1832DB
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   4320
      Picture         =   "frmWomensCasual.frx":183C41
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2400
      Picture         =   "frmWomensCasual.frx":1847F7
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   840
      Picture         =   "frmWomensCasual.frx":18533F
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "10"
      Height          =   255
      Left            =   5880
      TabIndex        =   31
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "9"
      Height          =   255
      Left            =   4080
      TabIndex        =   30
      Top             =   5640
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "7"
      Height          =   255
      Left            =   8520
      TabIndex        =   28
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   8520
      TabIndex        =   26
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmWomensCasual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmWomensCasual
'Author: Sean Johnson and Nick Lane
'Date Written: Sunday March 16th, 2007
'Objective of form: this is the women casual shoes form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Casual(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer

Option Explicit

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmWomensCasual.Hide
frmWomensShoes.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\CasualArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Casual(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Casual(j), Tab(25); Prices(j)
    End If
Next j
    If n > j Then   'this loops gives an error message of the user enters a number that doesnt correspond with the labeled items on the form
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


Private Sub cmdInfo1_Click()    'allows the user to view the specific information on the item
MsgBox "Brought back from the pavement of the past. The Nike Air Max 90 is a casual shoe with a retro running design that features Max Air® in the heel for comfort and cushioning. A leather and synthetic upper combines with a durable rubber outsole for comfort, support and durability.", , "Nike Women's Air Max 90 Premium"
End Sub

Private Sub cmdInfo10_Click()   'allows the user to view the specific information on the item
MsgBox "The Nike Air Sellwood is a crisp, clean-looking casual shoe inspired by the court. It features a leather and mesh upper with a visible air unit in the heel and a durable rubber outsole.", , "Nike Women's Air Sellwood"
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Max 90 women's casual shoe features a breathable mesh upper with an integrated saddle for midfoot support. The visible maximum Air-Sole® unit promises diving cushioning while the flex grooves on the outsole ensure comfort. A traction pattern on the outsole helps eliminate slip and slide.", , "Nike Women's Air Max 90"
End Sub
    
Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "Walk through the winter wonderlands in these deliciously warm boots by Nike. The Nike Winter High 2 has a suede upper and is lined with luxurious shearling to create a supercomfy boot. The front secures with a hidden zipper and pom-pom-enhanced laces. A rubber tread sole helps prevent spills on icy sidewalks.", , "Nike Women's Winter Hi 2"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "Stay cool and confident with inner Zen provided by the Nike Air Max '95 Zen Premium. Workout, walk and have fun in these cushioned shoes featuring a leather and mesh upper, a visible Max Air® unit in the heel and a rubber outsole.", , "Nike Women's Air Max '95 Zen Premium"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "Your feet will thank you for wrapping them in comfort and style while you wear the Nike Legend S/S. This casual kick features a leather upper and cupsole to keep feet well-cushioned and comfortable. The rubber outsole offers great traction and durability for long days of walking, shopping, or just hanging out.", , "Nike Women's Legend S/S"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "Add prestigious style to your casual shoe collection with the Nike Air Prestige II. Features of this shoe include a smooth leather upper for a super-soft feel, an Air® unit in the heel and a durable rubber outsole for a smooth ride.", , "Nike Women's Air Prestige II"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Blazer Low women's casual shoe is the original hoops shoe with a splash of pretty. Rich leather upper ensures a comfortable, smooth fit and feel. New patterns and fresh colors add hip style. Rubber outsole offers great traction.", , "Nike Women's Blazer Low"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "Originally a 1985 basketball shoe, the Nike Vandal puts a new twist on an old hoop classic. This casual street version, spawned from the original success of the Nike Vandal basketball shoe, features a plush leather upper with a VELCRO® brand fastener strap across the midfoot for a secure fit. Rubber outsole ensures great traction on the street or in the paint.", , "Nike Women's Vandal Low"
End Sub

Private Sub cmdInfo9_Click()    'allows the user to view the specific information on the item
MsgBox "A classic hoops shoe ready for all your downtown moves. The Nike Dunk Lo basketball shoe is styled for street wear. Leather/synthetic upper is comfortable and conforms to foot. EVA midsole offers stable, durable cushioning. Rubber outsole ensures great traction with a pivot point for quick turnarounds.", , "Nike Women's Dunk Low"
End Sub




