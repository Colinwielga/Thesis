VERSION 5.00
Begin VB.Form frmshop 
   Caption         =   "Shop"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form2"
   Picture         =   "frmshop.frx":0000
   ScaleHeight     =   8160
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check out!"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdmore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shop More!"
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7320
      Width           =   975
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   3600
      Picture         =   "frmshop.frx":1F8E9
      ScaleHeight     =   795
      ScaleWidth      =   675
      TabIndex        =   36
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H80000009&
      Height          =   615
      Left            =   3600
      Picture         =   "frmshop.frx":1FEA6
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   35
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   3600
      Picture         =   "frmshop.frx":203FB
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   34
      Top             =   5280
      Width           =   735
   End
   Begin VB.PictureBox Picture13 
      Height          =   615
      Left            =   3600
      Picture         =   "frmshop.frx":20959
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   33
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   3480
      Picture         =   "frmshop.frx":20E89
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   32
      Top             =   3240
      Width           =   855
   End
   Begin VB.PictureBox Picture11 
      Height          =   615
      Left            =   3480
      Picture         =   "frmshop.frx":213C9
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   31
      Top             =   2400
      Width           =   855
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   3480
      Picture         =   "frmshop.frx":21959
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   30
      Top             =   1320
      Width           =   855
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   3600
      Picture         =   "frmshop.frx":21FB7
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   29
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox Picture8 
      Height          =   735
      Left            =   480
      Picture         =   "frmshop.frx":223E8
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   28
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Picture7 
      Height          =   735
      Left            =   480
      Picture         =   "frmshop.frx":22810
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   27
      Top             =   6120
      Width           =   735
   End
   Begin VB.PictureBox Picture6 
      Height          =   735
      Left            =   480
      Picture         =   "frmshop.frx":22BA8
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   26
      Top             =   5280
      Width           =   735
   End
   Begin VB.PictureBox Picture5 
      Height          =   735
      Left            =   480
      Picture         =   "frmshop.frx":2312D
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   25
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox Picture4 
      Height          =   975
      Left            =   360
      Picture         =   "frmshop.frx":23668
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   24
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000009&
      Height          =   975
      Left            =   360
      Picture         =   "frmshop.frx":23D02
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   23
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   360
      Picture         =   "frmshop.frx":2454D
      ScaleHeight     =   795
      ScaleWidth      =   915
      TabIndex        =   22
      Top             =   1320
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   360
      Picture         =   "frmshop.frx":24B90
      ScaleHeight     =   813.333
      ScaleMode       =   0  'User
      ScaleWidth      =   915
      TabIndex        =   21
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox picoutput 
      Height          =   5655
      Left            =   6480
      Picture         =   "frmshop.frx":251BC
      ScaleHeight     =   5595
      ScaleWidth      =   4515
      TabIndex        =   20
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear"
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdtotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total"
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmd12 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmd13 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmd14 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmd15 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmd16 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmd9 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmd10 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd11 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmd8 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmd5 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add to Cart!"
      Height          =   615
      Left            =   1560
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   615
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to the main page"
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   1935
   End
End
Attribute VB_Name = "frmshop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Prize As Single

Private Sub cmd1_Click()
   Dim Guitar As Single
    Guitar = 229

    picoutput.Print "Badtz-Maru Fender Bronco Bass Guitar", FormatCurrency(Guitar)
    Prize = Prize + 229
End Sub

Private Sub cmd10_Click()
   Dim Plush3 As Single
    Plush3 = 6.5

    picoutput.Print "Kuromi Mascot Plush", FormatCurrency(Plush3)
    Prize = Prize + 6.5
End Sub

Private Sub cmd11_Click()
   Dim PCase As Single
    PCase = 6.5

    picoutput.Print "Little Twin Stars Pencil Case", FormatCurrency(PCase)
    Prize = Prize + 6.5
End Sub

Private Sub cmd12_Click()
   Dim Plush4 As Single
    Plush4 = 14

    picoutput.Print "My Melody Collectible Plush", FormatCurrency(Plush4)
    Prize = Prize + 14
End Sub

Private Sub cmd13_Click()
   Dim CCase As Single
    CCase = 7.5

    picoutput.Print "Pandapple Scented Candle in Case", FormatCurrency(CCase)
    Prize = Prize + 7.5
End Sub

Private Sub cmd14_Click()
   Dim Handbag2 As Single
    Handbag2 = 26

    picoutput.Print "Pankunchi Handbag", FormatCurrency(Handbag2)
    Prize = Prize + 26
End Sub

Private Sub cmd15_Click()
   Dim Pouch As Single
    Pouch = 11

    picoutput.Print "Pankunchi Pen Pouch", FormatCurrency(Pouch)
    Prize = Prize + 11
End Sub

Private Sub cmd16_Click()
   Dim Bag2 As Single
    Bag2 = 29

    picoutput.Print "Tenorikuma shoulder Bag", FormatCurrency(Bag2)
    Prize = Prize + 29
End Sub

Private Sub cmd2_Click()
   Dim Bag1 As Single
    Bag1 = 21

    picoutput.Print "Chi Chai Monchan Knapsack", FormatCurrency(Bag1)
    Prize = Prize + 21
End Sub

Private Sub cmd3_Click()
   Dim Container As Single
    Container = 9
    picoutput.Print "Chi Chai Monchan Lunch Container", FormatCurrency(Container)
    Prize = Prize + 9
End Sub

Private Sub cmd4_Click()

   Dim Cushion As Single
    Cushion = 20

    picoutput.Print "Chococat Knit Cushion", FormatCurrency(Cushion)
    Prize = Prize + 20
End Sub

Private Sub cmd5_Click()
   Dim Basket As Single
    Basket = 19

    picoutput.Print "Chococat Laundry Basket", FormatCurrency(Basket)
    Prize = Prize + 19
End Sub

Private Sub cmd6_Click()
   Dim Handbag As Single
    Handbag = 11

    picoutput.Print "Cinnamoroll Handbag", FormatCurrency(Handbag)
    Prize = Prize + 11
End Sub

Private Sub cmd7_Click()
   Dim Plush1 As Single
    Plush1 = 18

    picoutput.Print "Cinnamoroll Plush", FormatCurrency(Plush1)
    Prize = Prize + 18
End Sub

Private Sub cmd8_Click()
   Dim Key As Single
    Key = 4.5

    picoutput.Print "Deery-Lou Keycap", FormatCurrency(Key)
    Prize = Prize + 4.5
End Sub

Private Sub cmd9_Click()
   Dim Plush2 As Single
    Plush2 = 16

    picoutput.Print "Hello kitty Soft Plush", FormatCurrency(Plush2)
    Prize = Prize + 16
End Sub

Private Sub cmdback_Click()
frmmain.Visible = True
frmshop.Visible = False
End Sub

Private Sub cmdclear_Click()
    picoutput.Cls
    Prize = 0
End Sub

Private Sub cmdmore_Click()
MsgBox "Go to Sanrio.com see more items!"
frmweb.Show
frmshop.Hide
End Sub

Private Sub cmdout_Click()
frmshop.Hide
frmcheckout.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdtotal_Click()
    'declaring variables
    Dim Total As Single
    Dim Tax As Single
    'compute costs
    Tax = 0.06 * Prize
    Total = Tax + Prize
    'print results
    picoutput.Print "-------------------------------------"
    picoutput.Print "Subtotal", FormatCurrency(Prize)
    picoutput.Print "Tax", FormatCurrency(Tax)
    picoutput.Print "Total", FormatCurrency(Total)

End Sub
