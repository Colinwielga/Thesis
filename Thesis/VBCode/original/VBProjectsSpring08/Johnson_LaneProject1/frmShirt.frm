VERSION 5.00
Begin VB.Form frmShirt 
   Caption         =   "Shirts"
   ClientHeight    =   10560
   ClientLeft      =   1245
   ClientTop       =   630
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   Picture         =   "frmShirt.frx":0000
   ScaleHeight     =   10560
   ScaleWidth      =   12540
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   375
      Left            =   4680
      TabIndex        =   35
      Top             =   8400
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   2280
      ScaleHeight     =   1755
      ScaleWidth      =   6435
      TabIndex        =   34
      Top             =   6360
      Width           =   6495
   End
   Begin VB.CommandButton cmdBmen 
      Caption         =   "Back"
      Height          =   855
      Left            =   10680
      TabIndex        =   22
      Top             =   9600
      Width           =   1815
   End
   Begin VB.CommandButton cmdInfo11 
      Caption         =   "Info"
      Height          =   375
      Left            =   11160
      TabIndex        =   21
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo10 
      Caption         =   "Info"
      Height          =   375
      Left            =   9360
      TabIndex        =   20
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo9 
      Caption         =   "Info"
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   375
      Left            =   11160
      TabIndex        =   18
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   375
      Left            =   9240
      TabIndex        =   17
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   375
      Left            =   9480
      TabIndex        =   16
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton smdInfo5 
      Caption         =   "Info"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   375
      Left            =   11040
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   375
      Left            =   9240
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   375
      Left            =   11280
      TabIndex        =   12
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton smdInfo1 
      Caption         =   "Info"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   10080
      Width           =   1095
   End
   Begin VB.PictureBox Picture11 
      Height          =   1575
      Left            =   9000
      Picture         =   "frmShirt.frx":18004E
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   4920
      Width           =   1575
   End
   Begin VB.PictureBox Picture10 
      Height          =   1575
      Left            =   480
      Picture         =   "frmShirt.frx":180C19
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   6120
      Width           =   1575
   End
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   10920
      Picture         =   "frmShirt.frx":1817F2
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   10920
      Picture         =   "frmShirt.frx":1823B5
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   9000
      Picture         =   "frmShirt.frx":182E02
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   11040
      Picture         =   "frmShirt.frx":183782
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   8880
      Picture         =   "frmShirt.frx":1843E0
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   9120
      Picture         =   "frmShirt.frx":184D5B
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   7080
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   480
      Picture         =   "frmShirt.frx":1856FE
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   11040
      Picture         =   "frmShirt.frx":186BA8
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   7080
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   480
      Picture         =   "frmShirt.frx":1875F8
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   33
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   32
      Top             =   6240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   31
      Top             =   8520
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   8
      Left            =   8640
      TabIndex        =   30
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   7
      Left            =   8760
      TabIndex        =   29
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   6
      Left            =   8760
      TabIndex        =   28
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "10"
      Height          =   255
      Index           =   5
      Left            =   8880
      TabIndex        =   27
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "11"
      Height          =   255
      Index           =   4
      Left            =   10800
      TabIndex        =   26
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   3
      Left            =   10680
      TabIndex        =   25
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   2
      Left            =   10680
      TabIndex        =   24
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "9"
      Height          =   255
      Index           =   0
      Left            =   10800
      TabIndex        =   23
      Top             =   5040
      Width           =   135
   End
End
Attribute VB_Name = "frmShirt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmShirt
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the men's shirt form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Option Explicit
Dim Shirt(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Private Sub cmdBMen_Click()
'this button will hide this form and show the previous form
frmMenApparel.Show
frmShirt.Hide
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\shirtArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Shirt(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Shirt(j), Tab(25); Prices(j)
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

Private Sub cmdInfo10_Click()   'allows the user to view the specific information on the item
MsgBox "The Nike Pro Therma Extreme Long-Sleeve Mock is perfect as a first layer for wearing under athletic uniforms and training in extremely cold conditions. Dri-FIT® fabric is brushed on the inside for maximum warmth and comfort. Contrasting flat seam construction. Seams have been aligned below shoulder for enhanced mobility and less friction. Swoosh design trademark heat transfer at left front neck band. Nike FIT® trademark heat transfer at lower right. 90% polyester/10% spandex. Imported.", , "Nike Men's Pro Therma Extreme LS Mock"
End Sub
    
Private Sub cmdInfo11_Click()   'allows the user to view the specific information on the item
MsgBox "The Nike Big League Baseball long -leeve top features mid-sleeves in heavier Dri-FIT® fabric. Screenprinted Swoosh design trademark at the center chest and team jock tag at the lower left. Dri-FIT® 92% polyester/8% spandex with Dri-FIT® 90% polyester/10% spandex mid-sleeve. Imported.", , "Nike Men's L/S Big League Baseball Pro Top"
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Sphere Dry™ Roll Out polo features a self fabric collar and color insets across the front chest. Swoosh design trademark embroidered on the left collar. Authentic Nike label inserted on the left side seam. 100% polyester. Imported.", , "Nike Men's Sphere Dry Roll Out Polo"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Long-Sleeve Polo has a three-button placket, flat rib knit collar and vents at the side hem. Embroidered Swoosh design trademark at the upper right chest. 100% polyester. Imported.", , "Nike Men 's Long-Sleeve Polo"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "The Jordan Face Fleece takes style one step further with flashy metallic thread embroidery and hip screenprinted graphics. Made of 70% cotton/30% polyester, this full-zip fleece also has snap-close front pockets. Imported.", , "Jordan Lifestyle Men's Face Fleece"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Kobe 3 Tee captures the on-court blaze of the great player with the shine of glow-in-the-dark graphics! The raised-gel screenprinting appears on the front and left sleeve of this short-sleeve crew tee. Made of 100% cotton (5.7% organic). Imported.", , "Nike Men's Kobe 3 Tee"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike LBJ Logo Tee is a short-sleeve, crewneck T-shirt with a flocked L23 logo on the front. Swoosh design trademark at the upper back. 100% cotton (5.7% organic). Imported.", , "Nike Men's LBJ Logo Tee"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike short-sleeve Dri-FIT® mesh top features an embroidered arched wordmark on center chest as well as an embroidered Swoosh design trademark on center chest. Made from 100% polyester Dri-FIT®, a high-performance microfiber polyester fabric that actually pulls sweat away from the body and transports it to the fabric surface where it evaporates and leaves the skin cool and dry. It's all you need for hot days, and a critical base layer for cold days. Stay dry. Stay comfortable. No matter what. Imported.", , "Nike Men's Dri-Fit Mesh Top- North Carolina"
End Sub

Private Sub cmdInfo9_Click()    'allows the user to view the specific information on the item
MsgBox "The Jordan Since '85 Black Cat Tee is made of 100% cotton with screenprint on front and back. Front raised embroidery and contrast metallic chainstitch. Metallic thread embroidered Jumpman logo at bottom left hem and back neck. Imported", , "Jordan Lifestyle Men's Since 85 Black Cat Tee"
End Sub

Private Sub smdInfo1_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike LBJ Double-Double top is an engineer stripe short-sleeve polo with bias cut shoulder panel and a fully fashioned rib collar. Multi-color screenprint satin back neck tape. Satin embroidered L23 trademark and Swoosh design trademark. Embroidered lion's head on left chest. High-density screenprint and flock print all over the front and back. 100% cotton jersey with 97% cotton/3% spandex rib. Imported.", , "Nike Men's LBJ Double Double Polo"
End Sub

Private Sub smdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike City Squadron Tee is an athletic-fit short-sleeve tee. The front graphic treatment is a screenprint with a direct flock. 100% cotton. Imported.", , "Nike Men's City Squadron Tee"
End Sub
