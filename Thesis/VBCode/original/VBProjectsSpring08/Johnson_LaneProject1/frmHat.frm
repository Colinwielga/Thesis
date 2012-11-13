VERSION 5.00
Begin VB.Form frmHat 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hats"
   ClientHeight    =   9495
   ClientLeft      =   375
   ClientTop       =   555
   ClientWidth     =   12570
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmHat.frx":0000
   ScaleHeight     =   9495
   ScaleWidth      =   12570
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   2280
      ScaleHeight     =   1995
      ScaleWidth      =   8235
      TabIndex        =   20
      Top             =   6240
      Width           =   8295
   End
   Begin VB.CommandButton cmdBmen 
      Caption         =   "Back to Men's Apparel"
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdInfothree 
      Caption         =   "Info"
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfofour 
      Caption         =   "Info"
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfoseven 
      Caption         =   "Info"
      Height          =   375
      Left            =   9240
      TabIndex        =   16
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfofive 
      Caption         =   "Info"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   15
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfoeight 
      Caption         =   "Info"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfosix 
      Caption         =   "Info"
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfonine 
      Caption         =   "Info"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfotwo 
      Caption         =   "Info"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   480
      Picture         =   "frmHat.frx":47539
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   6960
      Picture         =   "frmHat.frx":4811C
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   4800
      Picture         =   "frmHat.frx":48BD8
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   7080
      Picture         =   "frmHat.frx":494E6
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   9360
      Picture         =   "frmHat.frx":4A175
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   480
      Picture         =   "frmHat.frx":4ABD6
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   9120
      Picture         =   "frmHat.frx":4B3D8
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2760
      Picture         =   "frmHat.frx":4BAC0
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   1095
      Left            =   10680
      TabIndex        =   1
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   600
      Picture         =   "frmHat.frx":4C5BC
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   29
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   28
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   7
      Left            =   9120
      TabIndex        =   27
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   25
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   4
      Left            =   8880
      TabIndex        =   24
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "9"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   5280
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   22
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "frmHat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmHat
'Author: Sean Johnson and Nick Lane
'Date Written: Saturday March 15th, 2007
'Objective of form: this is the men's hat form and from this form the
'                   user can find out information about the various items available,
'                   buy an item, and view the price of the item along with a running
'                   total of other items.

Option Explicit
Dim Hats(1 To 10) As String
Dim Prices(1 To 10) As Single
Dim ctr As Integer, j As Integer, c As String
Dim found As Boolean, n As Integer, i As Integer
Private Sub cmdBMen_Click()
'this button will hide this form and show the previous form
frmMenApparel.Show
frmHat.Hide
End Sub

'this button will allow the user to purchase an item
Private Sub cmdBuy_Click()

found = False

ctr = 0

'opens the array file to be read into a parallel array
Open App.Path & "\HatArray.txt" For Input As #1

'this loop will read the file into a parallel array which will be used when users selects an items
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Hats(ctr), Prices(ctr)
Loop
Close #1

'allows user to insert the number of the corresponding item he/she wishes to purchase or add to cart.
n = InputBox("Please enter the Number of the corresponding Hat You wish to purchase")

'this for loop will match the corresponding number above with the corresponding items in the array
For j = 1 To ctr
    If n = j Then
        found = True
        picResults.Print Hats(j), Tab(50); Prices(j)
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
        picResults.Print Tab(50); Tab(100); sum 'prints the users running total
    End If
         
Next i

End Sub



Private Sub cmdInfo_Click(Index As Integer) 'allows the user to view the specific information on the item
MsgBox "The Nike Lebron Las Vegas All Star Fitted Cap is made of 100% polyester with a 3D embroidered front logo. Imported.", , "Nike Lebron Vegas All Star Cap"
End Sub

Private Sub cmdInfofive_Click(Index As Integer) 'allows the user to view the specific information on the item
MsgBox "The Nike Perf Headphone Cap totally rocks, no matter what your genre of music. Plug in your music player, put on the cap, and you are geared up for whatever activity your day may hold. A great fit, fabulously comfortable and performance minded. Therma-FIT® fleece for the body and the interior headband. Therma-FIT® 96% polyester/4% spandex. Imported.", , "Nike Men's Performance Headphone Cap"
End Sub

Private Sub cmdInfotwo_Click()  'allows the user to view the specific information on the item
MsgBox "The Nike Classic College Swooshflex Cap is made of 97% cotton/3% spandex. This six-panel twill cap with a heavy buckram structure in front panels features a 3D team logo embroidered at center front. A Dri-FIT Swooshflex headband means it will comfortably fit most sizes.", , "Nike Men's Classic Swooshflex Cap- Ohio State"
End Sub

Private Sub cmdInfothree_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Classic College Swooshflex Cap is made of 97% cotton/3% spandex. This six-panel twill cap with a heavy buckram structure in front panels features a 3D team logo embroidered at center front. A Dri-FIT Swooshflex headband means it will comfortably fit most sizes.", , "Nike Men's Classic Swooshflex Cap- North Carolina"
End Sub

Private Sub cmdInfofour_Click() 'allows the user to view the specific information on the item
MsgBox "The Nike Classic College Swooshflex Cap is made of 97% cotton/3% spandex. This six-panel twill cap with a heavy buckram structure in front panels features a 3D team logo embroidered at center front. A Dri-FIT Swooshflex headband means it will comfortably fit most sizes. Swoosh design trademark at center back. Imported.", , "Nike Men's Classic Swooshflex Cap- Duke"
End Sub

Private Sub cmdInfosix_Click()  'allows the user to view the specific information on the item
MsgBox "Show your team spirit around town with this relaxed, six panel, washed cotton twill cap. The Nike Soccer Campus Cap features embroidered eyelets, an adjustable closure and a contrast underbill. Team name and badge logo embroidered at the center front. Swoosh design trademark embroidered at the center back. Small team logo woven label set into the back closure. 100% cotton. Imported.", , "Nike Soccer Campus Cap- Manchester United"
End Sub

Private Sub cmdInfoseven_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Perf Headphone Cap totally rocks, no matter what your genre of music. Plug in your music player, put on the cap, and you are geared up for whatever activity your day may hold. A great fit, fabulously comfortable and performance minded. Therma-FIT® fleece for the body and the interior headband. Therma-FIT® 96% polyester/4% spandex. Imported.", , "Nike Men's Performance Headphone Cap"
End Sub

Private Sub cmdInfoeight_Click()    'allows the user to view the specific information on the item
MsgBox "The Nike Tailwind Bones Cap takes the classic Tailwind and gives it a great update for the pinnacle in performance headwear. Made of 100% Dri-FIT® polyester, the style lines of this cap have been updated with open mesh for ventilation and contrast performance seam tape that is visible to the outside.", , "Nike Men's Tailwind Bones Cap"
End Sub

Private Sub cmdInfonine_Click() 'allows the user to view the specific information on the item
MsgBox "The new standard in performance headwear. The Nike Daybreak Cap is a lightweight running cap, weighing only 1.7 oz, with stream-lined, aerodynamic construction. Comfort bill and one-hand, quick-adjust closure ensures pinnacle performance.", , "Nike Men's Daybreak Cap"
End Sub

