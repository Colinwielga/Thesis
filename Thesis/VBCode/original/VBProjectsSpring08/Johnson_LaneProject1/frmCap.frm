VERSION 5.00
Begin VB.Form frmCap 
   Caption         =   "Capris"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   Picture         =   "frmCap.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FFFF&
      Caption         =   "Back"
      Height          =   495
      Left            =   5400
      TabIndex        =   26
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdBuy 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buy"
      Height          =   495
      Left            =   2640
      TabIndex        =   17
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   1935
      Left            =   2160
      ScaleHeight     =   1875
      ScaleWidth      =   6195
      TabIndex        =   16
      Top             =   2880
      Width           =   6255
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   495
      Left            =   720
      TabIndex        =   14
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   495
      Left            =   7080
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdInfo1 
      Caption         =   "Info"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.PictureBox Picture8 
      Height          =   1215
      Left            =   2760
      Picture         =   "frmCap.frx":16404E
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture7 
      Height          =   1215
      Left            =   4200
      Picture         =   "frmCap.frx":16450B
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox Picture6 
      Height          =   1215
      Left            =   600
      Picture         =   "frmCap.frx":164950
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      Height          =   1215
      Left            =   600
      Picture         =   "frmCap.frx":164E3F
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   4920
      Picture         =   "frmCap.frx":1653E9
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   600
      Picture         =   "frmCap.frx":1659EB
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   2400
      Picture         =   "frmCap.frx":165F0D
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   6960
      Picture         =   "frmCap.frx":16643A
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   6600
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "7"
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   6600
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   6600
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   6720
      TabIndex        =   21
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   600
      Width           =   135
   End
End
Attribute VB_Name = "frmCap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmCap
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the women capris shoes form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Capri(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Option Explicit



Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmCap.Hide
frmWomenApparel.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\CapriArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Capri(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Capri(j), Tab(25); Prices(j)
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

Private Sub cmdInfo1_Click()    'allows the user to view the specific information on the item
MsgBox "The perfect pant for all seasons. Relaxed-fit, stretch cotton spandex allows a custom fit while guaranteeing comfort with every wear. Style details include tonal stitching, unique seaming details, external drawcord and Nike logo embroidered at left hip. 96% cotton/4% spandex. Imported", , "MAX RELAX CAPRI PANTS-Blue Wave"

End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "It'll cater to your performance, yet satisfy your style. The Nike Women's Bamboo Practice Capri sports articulated knees, cinchable cuffs, and advanced ergonomic seaming for looking as good as it feels. A tie waist combines with a luxurious woven fabrication that includes earth-friendly bamboo and movement-friendly spandex. Swoosh design trademark embroidery at center back waist. Fabric: 56% cotton/41% bamboo/3% spandex poplin.", , "Bamboo Practice Capri"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "Jumpstart your dance moves. The Nike Rhythm Radiance Dance Women's Capri Pants sports a belt with sassy side ties for a combination of style and adjustability, while a floral outline embroidered on the right leg adds a feminine flair. Swoosh design trademark at left hip. Fabric: 98% cotton/2% spandex.", , "Nike Rhythm Radiance Women's Capri"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "When we say perfect, we mean it. The Nike Womens Perfect Fit Capri is sure to become a workout favorite after just one wear. This easy fit, straight leg capri with a subtle flare at the hem is constructed with sweat-wicking Dri-FIT fabric to help keep you cool and dry as you exercise. Wide waistband offers a low rise and comfortable fit. Swoosh design trademark at front waistband. FABRIC: Dri-FIT 289 g. 88% polyester/12% spandex jersey.", , "Perfect Fit Capri"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "Sophistication, ingenuity, and inspiration. The Women's Nike Guru Bamboo Yoga Capri features luxuriously soft, earth-friendly bamboo, along with an innovative, adaptable skirt that can be rolled up for a traditional pant look, delivering the ultimate in customizable coverage during your practice. Swoosh design trademark embroidered at left thigh. Fabric: 72% bamboo rayon/22% polyester/6% spandex plated jersey. Capri: Dri-FIT 65% bamboo rayon/27% polyester/8% spandex plated jersey.", , "Guru Bamboo Capri"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "Take your routine to the dance floor in the Nike Dance Woven Women's Capri Pants, featuring a capri cut for comfort and coverage, along with stretch fabrication and articulated knees for maximum mobility. Dri-FIT fabric wicks moisture and ventilates for cool-wearing comfort. Fabric: Body: Dri-FIT 100% polyester plain twill. Pocket/waist: Nike Sphere Dry 92% polyester/8% spandex circular knit mesh.", , "Nike Dance Woven Women's Capri Pants"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "Count on the Women's Nike Kaneel Capri to take you through many an outdoor escapade in style and comfort. Zip fly with snap closure. FABRIC: 100% polyester.", , "Kaneel Capri"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "The classic sweat to throw on before or after games, the Women's Nike Girl Graphic Capri can be worn rolled up or scrunched down for a unique look that's supremely comfy too. FABRIC: 80% cotton (5% organic)/20% polyester French terry.", , "Nike Girl Graphic Women's Capri"
End Sub
