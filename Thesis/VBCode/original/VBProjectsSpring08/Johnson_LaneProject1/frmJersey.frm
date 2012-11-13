VERSION 5.00
Begin VB.Form frmJersey 
   Caption         =   "Jerseys"
   ClientHeight    =   10065
   ClientLeft      =   270
   ClientTop       =   450
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   Picture         =   "frmJersey.frx":0000
   ScaleHeight     =   10065
   ScaleWidth      =   12900
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   495
      Left            =   6720
      TabIndex        =   26
      Top             =   7200
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   1935
      Left            =   4320
      ScaleHeight     =   1875
      ScaleWidth      =   6435
      TabIndex        =   25
      Top             =   5040
      Width           =   6495
   End
   Begin VB.CommandButton cmdBMen 
      Caption         =   "Back"
      Height          =   615
      Left            =   8640
      TabIndex        =   16
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Info"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Info"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Info"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   2640
      Picture         =   "frmJersey.frx":18004E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   6720
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   360
      Picture         =   "frmJersey.frx":180A22
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   4560
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   2640
      Picture         =   "frmJersey.frx":181479
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   2640
      Picture         =   "frmJersey.frx":181E88
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   360
      Picture         =   "frmJersey.frx":182A27
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   360
      Picture         =   "frmJersey.frx":183889
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   6720
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2640
      Picture         =   "frmJersey.frx":184481
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   360
      Picture         =   "frmJersey.frx":18509B
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   24
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   22
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   20
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   18
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmJersey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmJersey
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the men's jersey form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Jersey(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer

Private Sub cmdBMen_Click()
'this button will hide this form and show the previous form
frmMenApparel.Show
frmJersey.Hide
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()

found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\JerseyArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Jersey(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Jersey(j), Tab(25); Prices(j)
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

Private Sub Command1_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike USAB Vegas Twill Jersey is made from Dri-FIT® 100% polyester and features two-layer satin twill letters and numbers. Imported.", , "Nike Men's Vegas Twill Jersey-Lebron James"
End Sub

Private Sub Command2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Celtic Replica Jersey blends performance with style. 100% polyester Nike Sphere Dry™ with knit mesh panels to help you keep your cool during heated competition. Features Swoosh and team crest embroidery and a heat transfer sponsor on front chest. Imported.", , "Nike Men's Soccer Replica Jersey"
End Sub

Private Sub Command3_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Celtic Replica Jersey blends performance with style. 100% polyester Nike Sphere Dry™ with knit mesh panels to help you keep your cool during heated competition. Features Swoosh and team crest embroidery and a heat transfer sponsor on front chest. Imported.", , "Nike Men's Soccer Replica Jersey"
End Sub

Private Sub Command4_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike MLB Cooperstown Jersey is a respectful nod toward the former generations and how they made the national pastime what it is today. Made of 100% polyester, this full-button front baseball jersey features a tackle twill and embroidered team logo at left chest. Also features a team wordmark on right sleeve and Cooperstown jocktag at lower right hem, the team's home stadium screenprinted on center back neck, a Swoosh design trademark at right chest and a jocktag on lower left front. Imported.", , "Nike Men's Coopertown Jersey-Yankees"
End Sub

Private Sub Command5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Authentic Jersey is the authentic football jersey worn on the field. It features a 100% nylon front and back (side/cowl/sleeves vary by team) with tackle twill or screenprinted numbers on front, back and sleeves. Also includes school-specific silhouette, design lines, embellishments and fabrications. Imported.", , "Nike Men's College Football Authentic Jersey- Miami (Flor.)"
End Sub

Private Sub Command6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Authentic Jersey is the authentic football jersey worn on the field. It features a 100% nylon front and back (side/cowl/sleeves vary by team) with tackle twill or screenprinted numbers on front, back and sleeves. Also includes school-specific silhouette, design lines, embellishments and fabrications. Imported.", , "Nike Men's College Football Authentic Jersey- Texas"
End Sub

Private Sub Command7_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike MLB Cooperstown Jersey is a respectful nod toward the former generations and how they made the national pastime what it is today. Made of 100% polyester, this full-button front baseball jersey features a tackle twill and embroidered team logo at left chest. Also features a team wordmark on right sleeve and Cooperstown jocktag at lower right hem, the team's home stadium screenprinted on center back neck, a Swoosh design trademark at right chest and a jocktag on lower left front. Imported", , "Nike Men's Coopertown Jersey- St. Louis"
End Sub

Private Sub Command8_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Tradition Defined Reversible Jersey is a respectful nod at some of the greats. This throwback replica has team name and number sewn down tackle twill on the College Side. The USA Basketball Side features screenprinted team name and numbers. There is a Nike Swoosh or Jumpman trademark design on left chest and a Nike Team Sports jocktag applied to bottom hem for both sides. Made of 100% polyester. Imported.", , "Nike Men's TD Rev Twill Jrsy- North Carolina"
End Sub

