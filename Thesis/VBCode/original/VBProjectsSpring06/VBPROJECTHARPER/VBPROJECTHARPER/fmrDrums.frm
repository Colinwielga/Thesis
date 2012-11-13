VERSION 5.00
Begin VB.Form frmDrums 
   BackColor       =   &H80000006&
   Caption         =   "Buy drums"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCart 
      Caption         =   "Send Total to MY CART"
      Height          =   1935
      Left            =   6840
      TabIndex        =   16
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Home Page"
      Height          =   1935
      Left            =   8760
      TabIndex        =   15
      Top             =   9000
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Total"
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   10440
      Width           =   5775
   End
   Begin VB.CommandButton cmdComputeD 
      Caption         =   "Total Price for Drums"
      Height          =   1935
      Left            =   6840
      TabIndex        =   13
      Top             =   9000
      Width           =   1935
   End
   Begin VB.PictureBox picResultsdrums 
      Height          =   2295
      Left            =   840
      ScaleHeight     =   2235
      ScaleWidth      =   5715
      TabIndex        =   12
      Top             =   8160
      Width           =   5775
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000FF00&
      Caption         =   "By: Ben Harper"
      Height          =   255
      Left            =   9600
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000FF00&
      Caption         =   "Yamaha Electric Percussion $4,500"
      Height          =   855
      Left            =   8520
      TabIndex        =   11
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000FF00&
      Caption         =   "Yamaha Jam Session $1,279.99"
      Height          =   975
      Left            =   8520
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FF00&
      Caption         =   "5 pc. Yamaha Bush Series $850"
      Height          =   735
      Left            =   8520
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FF00&
      Caption         =   "5 pc. Yamaha kit $750"
      Height          =   975
      Left            =   8520
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "7 pc. Pearl Rodney Holmes Special Edition w/ DB $2,700"
      Height          =   1095
      Left            =   2400
      TabIndex        =   7
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "5 pc. Tama Old School Blu $1,400"
      Height          =   855
      Left            =   5520
      TabIndex        =   6
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "5 pc. Tama Fade Series $1,199.50"
      Height          =   735
      Left            =   5520
      TabIndex        =   5
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   "4 pc. Tama Fade Series $900"
      Height          =   975
      Left            =   5520
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "4 pc. Tama Starter Kit $650"
      Height          =   855
      Left            =   5520
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "5 pc Pearl Exports includes 2 Paiste Cymbals $869.99"
      Height          =   975
      Left            =   2400
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "7 pc. Pearl Exports w/ Double-Bass $1,999"
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "6 pc. Pearl Concert Series $1,250.99"
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image yamaEtc 
      Height          =   1710
      Left            =   6720
      Picture         =   "fmrDrums.frx":0000
      Top             =   6840
      Width           =   1800
   End
   Begin VB.Image yama3 
      Height          =   1545
      Left            =   6720
      Picture         =   "fmrDrums.frx":A092
      Top             =   5040
      Width           =   1800
   End
   Begin VB.Image Image13 
      Height          =   1800
      Left            =   6960
      Picture         =   "fmrDrums.frx":131AC
      Top             =   0
      Width           =   1380
   End
   Begin VB.Image tama4 
      Height          =   1425
      Left            =   3720
      Picture         =   "fmrDrums.frx":1B34E
      Top             =   6360
      Width           =   1800
   End
   Begin VB.Image yama2 
      Height          =   1275
      Left            =   6720
      Picture         =   "fmrDrums.frx":23928
      Top             =   3480
      Width           =   1800
   End
   Begin VB.Image yama1 
      Height          =   1335
      Left            =   6720
      Picture         =   "fmrDrums.frx":2B0F2
      Top             =   1920
      Width           =   1800
   End
   Begin VB.Image Image9 
      Height          =   780
      Left            =   3720
      Picture         =   "fmrDrums.frx":32E5C
      Top             =   120
      Width           =   1800
   End
   Begin VB.Image tama3 
      Height          =   1395
      Left            =   3720
      Picture         =   "fmrDrums.frx":377BE
      Top             =   4680
      Width           =   1800
   End
   Begin VB.Image tama2 
      Height          =   1380
      Left            =   3720
      Picture         =   "fmrDrums.frx":3FAC8
      Top             =   3000
      Width           =   1800
   End
   Begin VB.Image tama1 
      Height          =   1395
      Left            =   3720
      Picture         =   "fmrDrums.frx":47C6A
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Image pearl4 
      Height          =   1485
      Left            =   600
      Picture         =   "fmrDrums.frx":4FF74
      Top             =   6480
      Width           =   1800
   End
   Begin VB.Image pearl2 
      Height          =   1065
      Left            =   600
      Picture         =   "fmrDrums.frx":58AEE
      Top             =   3240
      Width           =   1800
   End
   Begin VB.Image pearl3 
      Height          =   1680
      Left            =   600
      Picture         =   "fmrDrums.frx":5EF08
      Top             =   4560
      Width           =   1800
   End
   Begin VB.Image pearl1 
      Height          =   1365
      Left            =   600
      Picture         =   "fmrDrums.frx":68CCA
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   600
      Picture         =   "fmrDrums.frx":70D04
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmDrums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Buy Drums Online (OnlineDrums.vbp)
'frmDrums (frmDrums)
'Ben Harper
'3/23/06
'This form is where the user can veiw pictures of each drum kit and general information
'such as prices, brand, and any accompanying features (cymbals, extra bass pedals).
'On this form, the total for drums alone can also be displayed before sending it to you cart.








Private Sub cmdCart_Click()
    frmCart.Visible = True  'sends drum total and drum name to cart
    frmDrums.Visible = False
    frmCart.picResults.Print "Drum Purchases", "Total"
    frmCart.picResults.Print "*************************************************"
    frmCart.picResults.Print "Your total for Drums is: ", FormatCurrency(Drumsum)
End Sub

Private Sub cmdClear_Click()  'clears total drum cost
picResultsdrums.Cls
Drumsum = 0
cmdComputeD.Visible = True
cmdCart.Visible = False

End Sub

Private Sub cmdComputeD_Click()  'computes total drum cost
    picResultsdrums.Print "*************************************"
    picResultsdrums.Print "Your Drum Total is: ", FormatCurrency(Drumsum)
    cmdComputeD.Visible = False
    cmdCart.Visible = True
End Sub

Private Sub cmdReturn_Click() 'returns to home page
frmDrums.Visible = False
frmHomePage.Visible = True
End Sub

Private Sub pearl1_Click()   'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "Pearl Concert Series ", FormatCurrency(1250.99, 2)
Drumsum = Drumsum + 1250.99
End Sub

Private Sub pearl2_Click()   'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "Pearl Exports w/ DB ", FormatCurrency(1999, 2)
Drumsum = Drumsum + 1999
End Sub

Private Sub pearl3_Click()   'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "Pearl Exports w/ Cymbals ", FormatCurrency(869.99, 2)
Drumsum = Drumsum + 869.99
End Sub

Private Sub pearl4_Click()   'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "Pearl R. Holmes SE w/ DB ", FormatCurrency(2700, 2)
Drumsum = Drumsum + 2700
End Sub

Private Sub tama1_Click()    'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "Tama Starter Kit    ", FormatCurrency(650, 2)
Drumsum = Drumsum + 650
End Sub

Private Sub tama2_Click()   'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "4 pc. Tama Fade Series ", FormatCurrency(900, 2)
Drumsum = Drumsum + 900
End Sub

Private Sub tama3_Click()   'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "5 pc. Tama Fade Series ", FormatCurrency(1199.5, 2)
Drumsum = Drumsum + 1199.5
End Sub

Private Sub tama4_Click()   'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "Tama Old School Blu Kit ", FormatCurrency(1400, 2)
Drumsum = Drumsum + 1400
End Sub

Private Sub yama1_Click()   'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "5 pc. Yamaha Kit   ", FormatCurrency(750, 2)
Drumsum = Drumsum + 750
End Sub

Private Sub yama2_Click()   'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "Yamaha Bush Series Kit ", FormatCurrency(850, 2)
Drumsum = Drumsum + 850
End Sub

Private Sub yama3_Click()    'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print "Yamaha Jam Session ", FormatCurrency(1279.99, 2)
Drumsum = Drumsum + 1279.99
End Sub

Private Sub yamaEtc_Click()    'adds drum cost to total drum cost and prints name of drums
picResultsdrums.Print " Yamaha Electric     ", FormatCurrency(4500, 2)
Drumsum = Drumsum + 4500
End Sub
