VERSION 5.00
Begin VB.Form frmWomensRunning 
   Caption         =   "Women's Running Shoes"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   Picture         =   "frmWomensRunning.frx":0000
   ScaleHeight     =   8700
   ScaleWidth      =   14895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   4680
      TabIndex        =   32
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdInfo10 
      Caption         =   "Info"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo9 
      Caption         =   "Info"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo8 
      Caption         =   "Info"
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo7 
      Caption         =   "Info"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo6 
      Caption         =   "Info"
      Height          =   255
      Left            =   11520
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo5 
      Caption         =   "Info"
      Height          =   255
      Left            =   9240
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo4 
      Caption         =   "Info"
      Height          =   255
      Left            =   7080
      TabIndex        =   15
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo3 
      Caption         =   "Info"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo2 
      Caption         =   "Info"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdInfo1 
      Caption         =   "Info"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1515
      ScaleWidth      =   6195
      TabIndex        =   10
      Top             =   6960
      Width           =   6255
   End
   Begin VB.PictureBox Picture10 
      Height          =   1455
      Left            =   480
      Picture         =   "frmWomensRunning.frx":18004E
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox Picture9 
      Height          =   1455
      Left            =   480
      Picture         =   "frmWomensRunning.frx":18092C
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   1455
      Left            =   8880
      Picture         =   "frmWomensRunning.frx":181450
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1455
      Left            =   2520
      Picture         =   "frmWomensRunning.frx":182008
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1455
      Left            =   4560
      Picture         =   "frmWomensRunning.frx":1829A4
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1455
      Left            =   11160
      Picture         =   "frmWomensRunning.frx":1832D0
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1455
      Left            =   6720
      Picture         =   "frmWomensRunning.frx":183B8A
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1455
      Left            =   4560
      Picture         =   "frmWomensRunning.frx":184514
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   2520
      Picture         =   "frmWomensRunning.frx":184FB6
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   360
      Picture         =   "frmWomensRunning.frx":185970
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "10"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "9"
      Height          =   255
      Left            =   4320
      TabIndex        =   30
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "7"
      Height          =   255
      Left            =   2280
      TabIndex        =   28
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "6"
      Height          =   255
      Left            =   10920
      TabIndex        =   27
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   8400
      TabIndex        =   26
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   6480
      TabIndex        =   25
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "frmWomensRunning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmWomensRunning
'Author: Sean Johnson and Nick Lane
'Date Written: Sunday March 16th, 2007
'Objective of form: this is the women running shoes form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Dim Running(1 To 20) As String
Dim Prices(1 To 20) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Option Explicit

Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmWomensRunning.Hide
frmWomensShoes.Show
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()
found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\RunningArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Running(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Shoe You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Running(j), Tab(25); Prices(j)
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
MsgBox "Let your feet feel great as you run in the Nike Air Zoom 360 II. This shoe features the most Air-Sole® cushioning ever engineered into a running shoe, which allows you to have a comfortable run. Designed with your comfort in mind, a breathable mesh upper a supportive rand and 360-degree reflectivity were added. Feel like you are barefoot with the unique full-length Air-Sole® unit which adds more durability and comfort. This shoe is great all the way down to the sole. Crafted with the durable BRS 1000™ rubber outsole and deep forefoot flex grooves your foot can maneuver around any terrain. Wt. 11.2 oz.", , "Nike Women's Air Max 360 II"
End Sub

Private Sub cmdInfo10_Click()   'allows the user to view the specific information on the item
MsgBox "The Nike Air Max Torch II running shoe is designed for the runner looking for great cushioning in a lightweight package. Mesh upper with synthetic rand adds breathability support. Visible Air Sole® uint offers plush cushioning in the heel. Phylon™ forefoot provides a great toe-off. Full-length BRS 1000™ carbon rubber outsole gives durability. Wt. 10.6 oz.", , "Nike Women's Air Max Torch II"
End Sub

Private Sub cmdInfo2_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Max Assail 5 is a cushioned, lightweight trail running shoe designed for multisurface use. Durable mesh upper with skeletal frame of support. Dual-fit system of webbing straps and synthetic overlays creates unsurpassed fit. Exterior heel counter overlay adds support of the foot on uneven terrain. Visible maximum Air-Sole® unit provides great cushioning. Full-length Phylon™ midsole offers lightweight flexibility. BRS 1000™ heel crash zone. Waffle® outsole blade traction technology allows for great traction on a variety of surfaces. Wt. 9.4 oz.", , "Nike Women's Air Max Assail V"
End Sub

Private Sub cmdInfo3_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Max Tailwind 2008 running shoe is made for the runner who needs big time durability and great cushioning. Open and breathable mesh upper with thin supportive overlays adds great fit. Full-length PU midsole gives unsurpassed durability. Maximum Air-Sole® units in heel and forefoot provide cushioning. BRS 1000™outsole offers great durability. Waffle® outsole design delivers great traction on a variety of surfaces. Wt. 11.4 oz.", , "Nike Women's Air Max Tailwind 2008"
End Sub

Private Sub cmdInfo4_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Air Max TL 4 running shoe is made for the runner seeking ride, durability and comfort. Lightweight synthetic mesh upper offers comfort, fit and breathability. A full-length Air-Sole® cushioning unit provides maximum impact protection. The BRS 1000™ carbon rubber outsole adds great traction in all conditions. Wt. 12.0 oz.", , "Nike Women's Air Max TL 4"
End Sub

Private Sub cmdInfo5_Click()    'allows the user to view the specific information on the item
MsgBox "The Shox Junga II is designed for the trail runner who needs all the technology Nike can offer in an off-road running shoe. Breathable and drainable mesh upper offers outdoor performance coupled with a dynamic FitFrame system for protection, support, fit and lock-down. Adaptive traction Nike Shox™ heel technology. forefoot Zoom Air™ unit, Phylon™ midsole forefoot pods. Waffle Fill™ outsole provides traction on hard surfaces and soft surfaces. Full-length carbon rubber outsole adds durability. Wt. 12.8 oz.", , "Nike Women's Shox Junga II"
End Sub

Private Sub cmdInfo6_Click()    'allows the user to view the specific information on the item
MsgBox "Crash, guide and contain is the mantra for the Nike Air Tri-D Run running shoe. Synthetic upper with TPU overlays offers great support. Soft density midsole foam cushions on heel strike, medium density foam guides the foot, firmer density foam contains the foot. BRS 1000™ carbon rubber Waffle® outsole ensures durability and outstanding traction. Wt. 11.4 oz.", , "Nike Women's Air Tri-D Run"
End Sub

Private Sub cmdInfo7_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Impax Tomahawk running shoe features a revolutionary new cushioning system in a lightweight leather profile. Frenetic design upper with perforated synthetic panel. One-piece leather tongue and vamp. Phylon™ midsole forefoot. Two-color rubber outsole with BRS 1000™ heel crash pad. Wt. 10.2 oz.", , "Nike Women's Impax Tomahawk"
End Sub

Private Sub cmdInfo8_Click()    'allows the user to view the specific information on the item
MsgBox "Fasten your seatbelts. The Nike Shox Turbo is back. Nike Shox™ columns absorb impact and return energy with every stride. Synthetic overlays lock the foot down for exceptional support. Transition wedge provides stabilizing linkage between Nike Shox™ columns and forefoot delivers an exceptionally smooth heel-to-toe transition. Strategically placed geometric deflection pods on the outsole offer great traction and a smooth ride.", , "Nike Women's Shox Turbo"
End Sub

Private Sub cmdInfo9_Click()    'allows the user to view the specific information on the item
MsgBox "With a coach or a specific training regime, the Nike Free 5.0 v3 running shoe can help to reduce injuries by strengthening your feet and legs. Unique, medial variable lacing system to customize a great fit. Engineered vents enhance the breathability. Phylite™ midsole is siped into an engineered pattern of flex grooves laterally and longitudinally. Strategic Waffle® outsole. BRS 1000™ rubber inserts are added to the Phylite™ to enhance durability.", , "Nike Women's Free 5.0 V3"
End Sub
