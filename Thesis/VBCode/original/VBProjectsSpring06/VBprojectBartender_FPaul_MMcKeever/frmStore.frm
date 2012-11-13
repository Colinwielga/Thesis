VERSION 5.00
Begin VB.Form frmStore 
   BackColor       =   &H0080C0FF&
   Caption         =   "Store"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Your Cart."
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdGuitar 
      Caption         =   "Add to Cart"
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txtGuitar 
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdSteins 
      Caption         =   "Add to Cart"
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox txtSteins 
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdHat 
      Caption         =   "Add to Cart"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox txtHat 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmdHelmet 
      Caption         =   "Add to Cart"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtHelmet 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute your Tax and Total"
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   5640
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   7920
      ScaleHeight     =   2715
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Head back to the bar!"
      Height          =   615
      Left            =   7920
      TabIndex        =   0
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label lblExperience 
      BackStyle       =   0  'Transparent
      Caption         =   "The Mike and Fred Bartending Experience:...Priceless"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   7200
      Width           =   8175
   End
   Begin VB.Label lblGuitar 
      BackStyle       =   0  'Transparent
      Caption         =   "Women love Guitars.  $300.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label lblSteins 
      BackStyle       =   0  'Transparent
      Caption         =   "Everyone will notice you at the bar when you use these mugs.  $13.75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6000
      TabIndex        =   10
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   1515
      Left            =   3120
      Picture         =   "frmStore.frx":0000
      Top             =   5280
      Width           =   1275
   End
   Begin VB.Image Image3 
      Height          =   2175
      Left            =   3120
      Picture         =   "frmStore.frx":0914
      Top             =   2640
      Width           =   2670
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fashion is always important when hitting the bar scene.  $19.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   7
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1350
      Left            =   360
      Picture         =   "frmStore.frx":255A
      Top             =   4680
      Width           =   1350
   End
   Begin VB.Label lblHelmet 
      BackStyle       =   0  'Transparent
      Caption         =   "Everyone needs a helmet now and then.  $23.50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   240
      Picture         =   "frmStore.frx":2A94
      Top             =   2640
      Width           =   1485
   End
   Begin VB.Label lblIntro 
      BackStyle       =   0  'Transparent
      Caption         =   "If you liked your bartending experience with Fred and Mike check out the rad merchandise we have!!!"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   9615
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Bartending School
'frmStore(Store)
'By Fred Paul & Michael McKeever
'March 22,2006
'The Store form advertises the products sold at the bar, and
'keeps inventory of each individual sale, and calculates subtotal
'and total with tax.

'Declare variables
    Dim hatamount, guitaramount, steinamount, helmetamount As Single
    Dim tax As Single
    Dim total As Single
    Dim subtotal As Single

Private Sub cmdBack_Click()
'This button returns the user to the bar form and hides the store
'from the user.
    frmStore.Hide
    frmBar.Show
End Sub

Private Sub cmdClear_Click()
'This button clears all information displayed in te piResults box
    picResults.Cls
End Sub

Private Sub cmdCompute_Click()
'This buttton computes the subtotal, tax, and total with tax
'when clicked/
    picResults.Print
    subtotal = hatamount + steinamount + guitaramount + helmetamount
        picResults.Print "Your subtotal is: "; FormatCurrency(subtotal)
        
    picResults.Print
    tax = 0.07 * subtotal
    picResults.Print "Tax:"; tax
    picResults.Print
    total = subtotal + tax
    picResults.Print "Total:"; FormatCurrency(total)
    
End Sub

Private Sub cmdGuitar_Click()
'This botton displays the subtotal amout for guitars in the
'picture box when clicked, it reads from input box.
    guitaramount = txtGuitar
    guitaramount = guitaramount * 300
    picResults.Print "Guitars: "; FormatCurrency(guitaramount)
End Sub

Private Sub cmdHat_Click()
'This displays the subtotal amount for hats in picResults
'when clicked, it reads from input box.
    hatamount = txtHat.Text
    hatamount = hatamount * 15.65
    picResults.Print "Cowboy Hats: "; FormatCurrency(hatamount)
End Sub

Private Sub cmdHelmet_Click()
'This botton displays the subtotal amout for helmets in the
'picture box when clicked, it reads from input box.
    helmetamount = txtHelmet.Text
    helmetamount = helmetamount * 23.5
    picResults.Print "Helmets: "; FormatCurrency(helmetamount)
End Sub

Private Sub cmdSteins_Click()
'This botton displays the subtotal amout for Steins in the
'picture box when clicked, it reads from input box.
    steinamount = txtSteins.Text
    steinamount = steinamount * 19.99
    picResults.Print "Steins: "; FormatCurrency(steinamount)
End Sub
