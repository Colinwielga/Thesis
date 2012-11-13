VERSION 5.00
Begin VB.Form frmSubtotal 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Subtotal"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   360
      Picture         =   "frmSubtotal.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   840
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdStartOver 
      Caption         =   "Start Over"
      Height          =   1095
      Left            =   840
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSubtotal 
      Caption         =   "Subtotal"
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080FF80&
      Height          =   5175
      Left            =   3360
      ScaleHeight     =   5115
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   2760
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Remember to Come Back and Find Great Deals                         Any Time of the Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Thanks For Shopping Electronics+!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "frmSubtotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjElectrPlus (Joe Dockendorff's VB Project.vbp)
'Form Name : frmSubtotal (frmSubtotal.frm)
'Author: Joe Dockendorff
'Date Written: March 13, 2004
'Purpose of Form: To get user to pick a product they would like
                 'to shop for and then give them a choice of models.
                 'The user can then pick the model of choice and
                 'add the product to their cart and checkout or
                 'shop around some more.  When the user is done, the
                 'total price is displayed.
                 
'Option Explicit is a command to force
'the user to declare all variables
'before they can be used.
Option Explicit

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStartOver_Click()
'This button takes the user back to the start form and allows them to start the whole
'process over.

frmSubtotal.Hide
frmStart.Show
End Sub

Private Sub cmdSubtotal_Click()
'This button displays the products selected and then calculates the subtota.
'It also asks the user if he wants to buy the products or not. If the user says
'"Y", then the user is asked to enter his name. If "N", then a thank you is displayed.
'With a "Y" answer, the Tax and total is calculated and displayed.
Dim Name As String

picResults.Cls

SubTotal = TVPrice(T) + CompPrice(C) + MP3Price(M)
Total = SubTotal + Tax

picResults.Print
picResults.Print
picResults.Print "Television:"
picResults.Print TV(T)
picResults.Print Tab(80); FormatCurrency(TVPrice(T))
picResults.Print
picResults.Print "Computer:"
picResults.Print Comp(C)
picResults.Print Tab(80); FormatCurrency(CompPrice(C))
picResults.Print
picResults.Print "MP3 Player:"
picResults.Print MP3(M)
picResults.Print Tab(80); FormatCurrency(MP3Price(M))
picResults.Print
picResults.Print "Sub-total:"; Tab(80); FormatCurrency(SubTotal)

Buy = InputBox("Would you like to buy these three models? Y or N", "Buy?")
If Buy = "Y" Then
    Name = InputBox("Please enter your name", "Name")
    Tax = 0.065 * SubTotal
    picResults.Print Tab(20); "Tax:"; Tab(80); FormatNumber(Tax, 2)
    picResults.Print Tab(80); "-------------------"
    picResults.Print
    picResults.Print Tab(20); "Total:"; Tab(80); FormatCurrency(Total)
    picResults.Print
    picResults.Print "Thank you for shopping with Electronics+ "; Name
    picResults.Print "We hope you choose us again!"
Else
    picResults.Print
    picResults.Print "Thank you for visiting Electronics+, to start over, click on the 'Start Over' button, else click 'Quit'"
    picResults.Print "We hope you come back again!"
End If


End Sub

Private Sub Command1_Click()
frmSubtotal.Hide
frmProdTVs.Show
End Sub

Private Sub Form_Load()

End Sub
