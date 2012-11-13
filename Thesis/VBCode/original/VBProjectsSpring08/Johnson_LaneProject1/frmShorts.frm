VERSION 5.00
Begin VB.Form frmShorts 
   Caption         =   "Shorts"
   ClientHeight    =   9675
   ClientLeft      =   1335
   ClientTop       =   990
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   Picture         =   "frmShorts.frx":0000
   ScaleHeight     =   9675
   ScaleWidth      =   12540
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   495
      Left            =   6720
      TabIndex        =   23
      Top             =   7680
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   1695
      Left            =   3240
      ScaleHeight     =   1635
      ScaleWidth      =   6915
      TabIndex        =   22
      Top             =   5760
      Width           =   6975
   End
   Begin VB.CommandButton cmdBmen 
      Caption         =   "Back"
      Height          =   615
      Left            =   3600
      TabIndex        =   14
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Info"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Info"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Info"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Info"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   840
      Picture         =   "frmShorts.frx":18004E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   6840
      Picture         =   "frmShorts.frx":18083A
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   6840
      Picture         =   "frmShorts.frx":18107C
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   3840
      Picture         =   "frmShorts.frx":181C0D
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   960
      Picture         =   "frmShorts.frx":1827EB
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   3840
      Picture         =   "frmShorts.frx":18334C
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   840
      Picture         =   "frmShorts.frx":1848B0
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   21
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   6600
      TabIndex        =   20
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   19
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   18
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   17
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   16
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   15
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "frmShorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmShorts
'Author: Sean Johnson and Nick Lane
'Date Written: Sunday March 16th, 2007
'Objective of form: this is the men's short form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim shorts(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Private Sub cmdBMen_Click()
'this button will hide this form and show the previous form
frmMenApparel.Show
frmShorts.Hide
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\shortsArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, shorts(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print shorts(j), Tab(25); Prices(j)
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

Private Sub Command1_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Dri-FIT® running short works overtime to keep you comfortable no matter what kind of paces you put it through. Elastic waist short with internal drawcord. On-seam pockets, contrast mesh insets at side leg. Side vent and self-fabric binding at hem. Swoosh design trademark embroidered at left hem. Dri-FIT® 86% polyester/14% spandex woven with Dri-FIT® 100% polyester crepe liner. 9 inch inseam. Imported.", , "Nike Men's 9 inch Dri-FIT Running Short"
End Sub

Private Sub Command2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Ohio State game short has an elastic waistband. Unique contrast inset design wrapping from the front to the back. Contrast Swoosh on left leg. 100% polyester double-knit flatback mesh with 100% nylon color insets. 9 inch inseam. Imported.", , "Nike Men's Ohio State Game Short"
End Sub
    
Private Sub Command3_Click()    'allows the user to view the specific information on the item
MsgBox "The NIke Twill Player Short is made of 100% polyester with a screenprinted sewn-down tackle twill appliqué. Imported.", , " Nike Men's Twill Shorts-Syracuse "
End Sub

Private Sub Command4_Click()   'allows the user to view the specific information on the item
MsgBox "The Reversible Nike Force mesh short is made with Dri-FIT fabric to keep moisture at bay, so you stay cool and dry all workout long. With team color on one side and contrast team color on the other side, you can match just about any tee with them. Featuring an embroidered team logo on team color side and an embroidered team name wordmark on contrast color side. Embroidered Swoosh design trademark on right leg of both sides. 100% polyester. Imported.", , "Nike Men's Reversible Mesh Shorts-UConn "
End Sub

Private Sub Command5_Click()    'allows the user to view the specific information on the item
MsgBox "The Reversible Nike Force mesh short is made with Dri-FIT fabric to keep moisture at bay, so you stay cool and dry all workout long. With team color on one side and contrast team color on the other side, you can match just about any tee with them. Featuring an embroidered team logo on team color side and an embroidered team name wordmark on contrast color side. Embroidered Swoosh design trademark on right leg of both sides. 100% polyester. Imported.", , "Nike Men's Reversible Mesh Shorts-Arizona"
End Sub

Private Sub Command6_Click()    'allows the user to view the specific information on the item
MsgBox "Get the crowd's attention with the Nike Miami Game Short. Elastic waistband and inside drawcord offer the perfect fit no matter how fast you move. Contrast side inset and Swoosh design trademark add visual appeal. Dri-FIT® 100% polyester jersey. 9 inch inseam. Imported.", , "Nike Men's Stock Hurricane Short"
End Sub

Private Sub Command7_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Tradition Defined Shorts are an excellent throwback option for both on the court and off. These longer length game replica shorts are made of 100% polyester and feature a 13 inch inseam. Screenprinted, sewn-down tackle twill applications as well as a Swoosh on upper left leg add to the authenticity. Imported.", , "Nike Men's Traditon Defined Shorts- North Carolina"
End Sub
