VERSION 5.00
Begin VB.Form frmgiftshop 
   BackColor       =   &H000000FF&
   Caption         =   "SJU HOCKEY ONLINE GIFT SHOP"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form3"
   ScaleHeight     =   10770
   ScaleWidth      =   15240
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Total and Keep Shopping"
      Height          =   975
      Left            =   10320
      TabIndex        =   21
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Thank You, Come Again"
      Height          =   975
      Left            =   8400
      TabIndex        =   10
      Top             =   9120
      Width           =   1695
   End
   Begin VB.TextBox txtj 
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate Your Total"
      Height          =   975
      Left            =   6480
      TabIndex        =   8
      Top             =   9120
      Width           =   1815
   End
   Begin VB.TextBox txtmug 
      Height          =   735
      Left            =   12720
      TabIndex        =   7
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txthooded 
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   9240
      Width           =   735
   End
   Begin VB.TextBox txtk 
      Height          =   615
      Left            =   13680
      TabIndex        =   5
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txth 
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txts 
      Height          =   615
      Left            =   12120
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtm 
      Height          =   615
      Left            =   8400
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtb 
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FF0000&
      Height          =   4935
      Left            =   7320
      ScaleHeight     =   4875
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label lbldiscount 
      BackColor       =   &H00FF0000&
      Caption         =   "20 % Off Your First ONLINE Purchase!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label lblwelcom 
      BackColor       =   &H00FF0000&
      Caption         =   $"Form3.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   20
      Top             =   840
      Width           =   6495
   End
   Begin VB.Label lblmug 
      BackColor       =   &H00FF0000&
      Caption         =   "Beer Mug................................$15 Doubles as a Milk Mug"
      Height          =   495
      Left            =   12000
      TabIndex        =   19
      Top             =   8880
      Width           =   2415
   End
   Begin VB.Label lblsweats 
      BackColor       =   &H00FF0000&
      Caption         =   "Hooded Sweatshirts.............$35 Stay warm in those freezing ice rinks."
      Height          =   615
      Left            =   3720
      TabIndex        =   18
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label lblkeyring 
      BackColor       =   &H00FF0000&
      Caption         =   "Skate Key-Chain.................$8  Take a piece of the game with you wherever you go."
      Height          =   615
      Left            =   12960
      TabIndex        =   17
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label lblhat 
      BackColor       =   &H00FF0000&
      Caption         =   "Baseball Cap........................$20 One Size Fits All"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lblseatcushion 
      BackColor       =   &H00FF0000&
      Caption         =   "Cushion..................................$5 Give your butt a break."
      Height          =   495
      Left            =   11280
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblminni 
      BackColor       =   &H00FF0000&
      Caption         =   "Mini-hockey set ................$20 The're Going Fast!!!!"
      Height          =   495
      Left            =   7680
      TabIndex        =   14
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblbanner 
      BackColor       =   &H00FF0000&
      Caption         =   "SJU Banner ......................  $50 Place your order below."
      Height          =   615
      Left            =   3720
      TabIndex        =   13
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label lbljersey 
      BackColor       =   &H00FF0000&
      Caption         =   "Hockey Jersey..................$75 How many do you want?"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   9240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "By: Matthew Czech"
      Height          =   255
      Left            =   9840
      TabIndex        =   11
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Image immug 
      Height          =   2250
      Left            =   12120
      Picture         =   "Form3.frx":00E5
      Top             =   6600
      Width           =   2115
   End
   Begin VB.Image imgsweats 
      Height          =   2190
      Left            =   3720
      Picture         =   "Form3.frx":4891
      Top             =   6240
      Width           =   2250
   End
   Begin VB.Image imgmini 
      Height          =   1980
      Left            =   7680
      Picture         =   "Form3.frx":A8DF
      Top             =   360
      Width           =   2100
   End
   Begin VB.Image imgkeyring 
      Height          =   2175
      Left            =   13320
      Picture         =   "Form3.frx":C501
      Top             =   2760
      Width           =   1590
   End
   Begin VB.Image imghat 
      Height          =   1755
      Left            =   360
      Picture         =   "Form3.frx":DFB9
      Top             =   3000
      Width           =   2250
   End
   Begin VB.Image imgseet 
      Height          =   2250
      Left            =   11280
      Picture         =   "Form3.frx":FCCF
      Top             =   240
      Width           =   2250
   End
   Begin VB.Image imgbanner 
      Height          =   2250
      Left            =   3720
      Picture         =   "Form3.frx":11EEA
      Top             =   2640
      Width           =   2250
   End
   Begin VB.Image imgjersey 
      Height          =   2520
      Left            =   480
      Picture         =   "Form3.frx":142E8
      Top             =   6720
      Width           =   2250
   End
End
Attribute VB_Name = "frmgiftshop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Project Name : SJU HOCKEY (Matthew Czech's VB Project.vbp)
'Form Name : frmgiftshop(Form3.frm)
'Author: Matthew Czech
'Date Written: March 12, 2003
'Purpose of form: to offer mechandise and to calculate the cost of
            'different selected items.



Dim j As Single, b As Single, m As Single, s As Single, h As Single, k As Single, hooded As Single, mug As Single

Private Sub Command1_Click()

j = Val(txtj.Text)  'Allows 0 to be assumed if space is not filled out
b = Val(txtb.Text)
m = Val(txtm.Text)
s = Val(txts.Text)
h = Val(txth.Text)
k = Val(txtk.Text)
hooded = Val(txthooded.Text)
mug = Val(txtmug.Text)
SubTotal = (j * 75) + (b * 50) + (m * 20) + (s * 5) + (h * 20) + (k * 8) + (hooded * 35) + (mug * 15) 'calculates total
If j > 0 Then
    picresults.Print "Hockey Jersey"; j; "* 75", , FormatCurrency(j * 75) 'prints selected items
    picresults.Print
End If
If b > 0 Then
    picresults.Print "SJU Banner"; b; "* 50", , FormatCurrency(b * 50)
    picresults.Print
End If
If m > 0 Then
    picresults.Print "Mini-Sticks"; m; "* 20", , FormatCurrency(m * 20)
    picresults.Print
End If
If s > 0 Then
    picresults.Print "Cushion"; s; "* 5", , , FormatCurrency(s * 5)
    picresults.Print
End If
If h > 0 Then
    picresults.Print "Hat"; h; "* 20", , , FormatCurrency(h * 20)
    picresults.Print
End If
If k > 0 Then
    picresults.Print "Key-Chain"; k; "* 8", , FormatCurrency(k * 8)
    picresults.Print
End If
If hooded > 0 Then
    picresults.Print "Sweat-Shirt"; hooded; "* 35", , FormatCurrency(hooded * 35)
    picresults.Print
End If
If mug > 0 Then picresults.Print "Beer Mug"; mug; "* 15", , FormatCurrency(mug * 15)
picresults.Print "+++++++++++++++++++++++++++++++++++++++++++++++"
picresults.Print
picresults.Print "SubTotal", , , FormatCurrency(SubTotal)

discount = SubTotal * 0.8
picresults.Print "20% discount price", , FormatCurrency(discount)

tax = discount * 0.065
picresults.Print "Tax", , , FormatCurrency(tax)
picresults.Print "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
picresults.Print
Total = discount + tax
picresults.Print "Total", , , FormatCurrency(Total)
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
picresults.Cls
txtj.Text = 0   'sets blank spaces equal to zero items wanted
txtb.Text = 0
txtm.Text = 0
txts.Text = 0
txth.Text = 0
txtk.Text = 0
txthooded.Text = 0
txtmug.Text = 0
End Sub

