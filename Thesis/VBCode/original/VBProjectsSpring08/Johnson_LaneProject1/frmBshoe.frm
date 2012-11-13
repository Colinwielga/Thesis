VERSION 5.00
Begin VB.Form frmBshoe 
   Caption         =   "Basketball Shoes"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   Picture         =   "frmBshoe.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuy 
      BackColor       =   &H8000000D&
      Caption         =   "Buy"
      Height          =   495
      Left            =   8040
      Picture         =   "frmBshoe.frx":18004E
      TabIndex        =   32
      Top             =   7680
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   2175
      Left            =   720
      ScaleHeight     =   2115
      ScaleWidth      =   7155
      TabIndex        =   31
      Top             =   6600
      Width           =   7215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H8000000D&
      Caption         =   "Back"
      Height          =   495
      Left            =   8040
      Picture         =   "frmBshoe.frx":184A60
      TabIndex        =   20
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdinfo10 
      Caption         =   "Info"
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdinfo9 
      Caption         =   "Info"
      Height          =   255
      Left            =   5880
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdinfo8 
      Caption         =   "Info"
      Height          =   255
      Left            =   9120
      TabIndex        =   17
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdinfo7 
      Caption         =   "Info"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdinfo6 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdinfo5 
      Caption         =   "Info"
      Height          =   255
      Left            =   10680
      TabIndex        =   14
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdinfo4 
      Caption         =   "Info"
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdinfo3 
      Caption         =   "Info"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdinfo2 
      Caption         =   "Info"
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox Picture10 
      Height          =   1575
      Left            =   8280
      Picture         =   "frmBshoe.frx":189472
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   9000
      Picture         =   "frmBshoe.frx":18A00D
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   8160
      Picture         =   "frmBshoe.frx":18AC91
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   5640
      Picture         =   "frmBshoe.frx":18B649
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   360
      Picture         =   "frmBshoe.frx":18C1A3
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   10560
      Picture         =   "frmBshoe.frx":18CA7F
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   3000
      Picture         =   "frmBshoe.frx":18D341
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   2760
      Picture         =   "frmBshoe.frx":18DE71
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   5640
      Picture         =   "frmBshoe.frx":18E938
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   480
      Picture         =   "frmBshoe.frx":18F43C
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   10
      Left            =   8760
      TabIndex        =   30
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "9"
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   29
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "10"
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   28
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   27
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   26
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   25
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   24
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   3
      Left            =   10320
      TabIndex        =   23
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "frmBshoe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmBshoe
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the men's basketball shoes form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Option Explicit
Dim Bshoes(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmBshoe.Hide
frmShoes.Show
End Sub


'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()

found = False
ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\BshoeArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Bshoes(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Bshoes(j), Tab(25); Prices(j) 'prints the item and price so that the user can see the price of the individual item
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
        picResults.Print Tab(25); Tab(50); sum 'prints the users running total
    End If
         
Next i

End Sub

Private Sub cmdInfo_Click()     'allows the user to view the specific information on the item
MsgBox "A mixture of the best technology from the past, present and future come together to create the world's ultimate basketball shoe. Featured in colors that Kobe himself has worn on the court, this shoe is engineered to propel the game’s best to the highest level of explosiveness, ride and quickness. The Nike Air Zoom Huarache 2K4 Laser features a rich, ultra-thin, lightweight leather and suede upper with intricate laser etching details that conforms to the foot. The external heel counter offers enhanced stability and heel fit, while the integrated ankle strap locks the foot over the footbed. The hyper lightweight Phylon™ midsole with heel and forefoot Zoom Air units provides a low-to-the-ground responsive ride and the solid rubber outsole with maximum traction herringbone pattern, carbon fiber shank plate adds torsional rigidity and spring. Wt. 14.4 oz.", , "Nike Men's Air Zoom Huarache 2K4 (Laser)"
End Sub

Private Sub cmdInfo10_Click()   'allows the user to view the specific information on the item
MsgBox "The Nike Air Zoom Huarache Elite TB is the perfect combination of lightweight comfort, responsiveness and lateral support. A futuristic and modern Uptempo basketball shoe for the game's quickest and most nimble players. Durable and rich synthetic leather upper. Ankle lockdown strap adds enhanced support. Neoprene ankle inserts offer breathable and comfortable fit. Seamless internal bootie. Lightweight Phylon™ midsole combined with elements of the Free Technology in the forefoot for enhanced flexibility and ride. Zoom Air™ in the heel and forefoot. TPU plate in the heel gives heel lockdown. Durable solid rubber outsole with maximum traction herringbone pattern. Wt. 15.6 oz.", , "Nike Men's Air Zoom Huarache Elite II TB "
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "Nike celebrates 25 years of Air Force with this modern, plush, sophisticated, responsive Nike Air Force 25 basketball shoe that evolves the heritage of the franchise. Seamless, full-grain leather upper with molded panels. Full-length innersleeve with Nike Sphere™ lining. Internal dual-pull lacing system. Phylon™ cupsole construction combines a full-length Zoom Air™ unit with maximum-volume Air-Sole® unit in heel. Carbon fiber midfoot shank. Clear rubber outsole wraps upward, intersecting the midsole. Solid rubber pivot circles with inset herringbone pattern in the heel and forefoot.", , "Nike Men's Air Force 25 League Pack CC"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "With the Jordan Collezione 14/9, it's all about the celebration of MJ's legendary number 23. The combination of the Air Jordan Retro XIV and the Air Jordan Retro IX creates the infamous number 23, and also gives you double the luxe style. The Air Jordan Retro XIV is a high-performance, luxury basketball shoe for the basketball player in the driver's seat. The leather upper provides comfort and support while a unique internal lacing system enhances the fit. The Phylon™ midsole with Zoom Air® unit in the heel and forefoot provides superior cushioning while maintaining a lightweight feel. The solid rubber outsole with herringbone tread pattern delivers maximum traction, and the composite shank plate adds support and stability.", , "Jordan Men's Collezione 14/9"

End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "Jordan and Air Force 1 have come together to release the Jordan AJF 12 LS basketball shoe. The blend of iconic Air Jordan and timeless Air Force 1 creates a whole new definition of sleek and enviable performance hoops shoe style, showcased by the full-grain leather and nubuck upper that is combined with a classic, removable ankle strap. The midsole is wrapped in leather and combined with an internal Air-Sole® unit while the outsole features a circular traction pattern. This official collaboration between Jordan and Air Force 1 is sure to make heads turn for years to come.", , "Jordan Men's AJF 12 LS"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "Elevate your basketball performance to a whole new level with the Jordan Team Elite 08 basketball shoe. This high flyin' Jordan shoe features a combination of floater and rich nubuck leathers with a foam-backed collar and perforated, synthetic leather tongue. The midsole is lightweight Phylon™ with a double-stacked Zoom Air in the heel for the ultimate reaction. A solid rubber outsole with herringbone traction pods add a great grip on the court.", , "Jordan Men's Team Elite 08"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Zoom Kobe III basketball shoe sports a full-grain leather upper. Full-length fit sleeve with a full-length Zoom Air™ unit sockliner. New Natural Motion Phylon™ midsole allows the best transition and court feel. Carbon fiber midfoot shank. Natural Motion pods with herringbone pattern provide unprecedented traction and court feel. Wt. 16.4 oz.", , "Nike Men's Kobe III"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "Engineered specifically for LeBron James, the Nike Zoom LeBron V basketball shoe represents the journey of the king that provides the ultimate lightweight protection. Phyposite bucket with plush leather overlays. Dynamic Fit innersleeve adds royal comfort. Integrated strap offers lockdown. Full-length Zoom Air™ unit delivers the ultimate responsive ride with double stacked Zoom Air™ in the heel. Carbon fiber spring plate. Optimal motion flex grooves with a clear rubber outsole to give enhanced court traction. Wt. 19.2 oz.", , "Nike Men's Zoom LeBron V"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "A shoe built with the future of basketball in mind. The Nike Zoom LeBron IV basketball shoe was created for the next generation of player, LeBron James. Foamposite technology in an integrated upper/midsole provides seamless comfort and support. Slow recovery memory foam in the upper collar provides a balance of flexibility and support with a strap for added lock down. This is a full-length Zoom Air™ unit for the ultimate responsive ride with a full-length TPU plate that provides torsional rigidity for quick moves. Clear outsole with solid rubber herringbone pods delivers on-court traction. Wt. 20.7 oz.", , "Nike Men's Zoom LeBron IV"
End Sub

Private Sub cmdInfo9_Click()    'allows the user to view the specific information on the item
MsgBox "The Jordan Melo M4 is a custom-built shoe created for all-star performance and design for the specific needs of Carmelo Anthony's game. The upper is a combination of full-grain leathers and rich suede nubucks for a superior and lightweight performance. A customizable lockdown lacing system and ankle sleeve supply an excellent fit. The Foamposite is the heel provides additional support and stability, while the Neoprene sock system provides a custom fit for the collar. There is a Phylon™ midsole with a full-length Nike Air Cushioning system and shank system that maximizes support. The solid rubber outsole has a modified herringbone shape and offers performance traction. Separated heel and forefoot pods supply maximum agility and the lightweight shank plate offers an engineered support. Wt. 14.4 oz.", , "Jordan Men's Melo M4"
End Sub


