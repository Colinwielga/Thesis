VERSION 5.00
Begin VB.Form frmdresses 
   Caption         =   "Dresses/Skirts"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   Picture         =   "frmdresses.frx":0000
   ScaleHeight     =   8085
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   5640
      TabIndex        =   20
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   4680
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   3000
      ScaleHeight     =   1995
      ScaleWidth      =   4395
      TabIndex        =   12
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   495
      Left            =   7800
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   495
      Left            =   8040
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo1 
      Caption         =   "Info"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox Picture6 
      Height          =   1215
      Left            =   7800
      Picture         =   "frmdresses.frx":18B37
      ScaleHeight     =   1155
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.PictureBox Picture5 
      Height          =   1455
      Left            =   6360
      Picture         =   "frmdresses.frx":1916C
      ScaleHeight     =   1395
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox Picture4 
      Height          =   1455
      Left            =   8040
      Picture         =   "frmdresses.frx":194E0
      ScaleHeight     =   1395
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Height          =   1335
      Left            =   4080
      Picture         =   "frmdresses.frx":19A65
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   7680
      Picture         =   "frmdresses.frx":1A535
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   1560
      Picture         =   "frmdresses.frx":1AB20
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   9000
      TabIndex        =   19
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   9000
      TabIndex        =   18
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   9000
      TabIndex        =   17
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmdresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmDresses
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the women dresses shoes form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Dress(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer

Option Explicit

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmdresses.Hide
frmWomenApparel.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\DressArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Dress(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Dress(j), Tab(25); Prices(j)
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

Private Sub cmdInfo1_Click() 'allows the user to view the specific information on the item
MsgBox "Shake up your running routine with Nike's sassy Adventure women's skirt. A key trend for the fashion-conscious consumer, this skirt is a versatile option if you're on the go. A knit skirt with an internal compression short allows for excellent mobility, coverage and comfort. A back origami pocket provides easy storage for keys, cards or media player and reflective piping for added visibility and safety. Swoosh design trademark embroidered at left hem. Imported.", , "Nike Women's Adventure Skirt"
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "Double your style options with the Women's Nike Heritage Pleated Skirt. This reversible skirt is plaid on one side, solid on the other, but with its breathable Dri-FIT fabric, its performance is solid all around. ", , "Heritage Pleated Skirt"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "When the heat is on, the Nike Women's Tennis Day Dress is certain to keep you cool and ready to control the court. Designed with comfortable Dri-FIT fabric, a built-in bra and a chic silhouette for wowing the crowd and dominating your opponent. ", , "Nike Women's Tennis Day Dress"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "A look loaded with style and power, the Serena Women's Tennis Day Dress features all the best to help you play just like Williams. This dress is loaded with features that you'll be grateful for on the court, like sweat-wicking Dri-FIT fabric, Nike Sphere Dry mesh, comfortable straps, a built-in bra and a silhouette that's made to move. ", , "Serena Women's Tennis Day Dress "
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "Pretty pleats meet high-performance control. The Nike Women's Control Pleated Tennis Skirt features Dri-FIT mid-weight ray that wicks away moisture for cooling, while pleats provide a fun flair and room to move. A set-on contrast stripe cascades around the body, while an internal short provides coverage and compression. ", , "Control Pleated Skirt "
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "Bold, contrasting leaf patterns distinguish the Women's Nike Print Border Skirt from other more mundane tennis attire. Features a contrasting border down the sides and along hem. ", , "Print Border Skirt "
End Sub
