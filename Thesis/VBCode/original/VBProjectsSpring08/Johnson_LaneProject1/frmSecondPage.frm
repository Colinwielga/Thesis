VERSION 5.00
Begin VB.Form frmSecondPage 
   Caption         =   "Men or Women Section"
   ClientHeight    =   9525
   ClientLeft      =   915
   ClientTop       =   660
   ClientWidth     =   13305
   BeginProperty Font 
      Name            =   "Curlz MT"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmSecondPage.frx":0000
   ScaleHeight     =   9525
   ScaleWidth      =   13305
   Begin VB.CommandButton cmdCheckout 
      Caption         =   "Checkout"
      BeginProperty Font 
         Name            =   "Charlemagne Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Picture         =   "frmSecondPage.frx":4A12
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      DownPicture     =   "frmSecondPage.frx":9424
      BeginProperty Font 
         Name            =   "Charlemagne Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Picture         =   "frmSecondPage.frx":DE36
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdTrivia 
      Caption         =   "Trivia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7200
      Picture         =   "frmSecondPage.frx":12848
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1800
      Picture         =   "frmSecondPage.frx":19B0B
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdFemaleSection 
      Caption         =   "Women"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   6480
      MaskColor       =   &H008080FF&
      Picture         =   "frmSecondPage.frx":19F1F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton cmdMaleSection 
      Caption         =   "Men"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      Picture         =   "frmSecondPage.frx":1E875
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Label lblTrivia 
      BackColor       =   &H80000006&
      Caption         =   "            Take the Nike Trivia Challenge!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   8400
      Width           =   3375
   End
End
Attribute VB_Name = "frmSecondPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmSecondPage
'Author: Sean Johnson and Nick Lane
'Date Written: Friday March 14th, 2007
'Objective of form: this form allows the user to enter different sections of the store.
'                   from here the user can go into the men or women section, or choose
'                   to complete the trivia challenges set up for extra discounts on purchases.


Private Sub cmdBack_Click()
'this button will hide this form and show the previous form
frmSecondPage.Hide
frmFrontPage.Show
End Sub

Private Sub cmdCheckout_Click()
Dim finaltotal As Single

MsgBox "Thankyou for shopping at Nike Town", , "THANKYOU" 'this message thanks the user for shopping in nike town
MsgBox "Your final total " & sum, , "Total" 'shows the total before discount is taken off
finaltotal = (sum * discount) ' calculates the discount amount
sum = sum - finaltotal 'calculates the final total after discount is accounted for
MsgBox "After your discount earned from trivia, your final total is " & FormatCurrency(sum), , "Bill" 'displays the final total


 If sum = 0 Then ' this if loop lets the computer go directly to the font page if no items are purchased
    frmFrontPage.Show
    frmSecondPage.Hide
End If
    If sum > 0 Then 'allows the user to decide whether they will pay with cash or credit if an item or items are bought
        c = InputBox("Will you be paying Cash of Credit?", , "Cash Or Credit")
            'gives the customer instructions if the credit option of payment is choosen
            If c = "Credit" Then
                MsgBox "Proceed to the agent in aisle 4"
             End If
    End If
    
    'gives the customer instructions if the cash option of payment is choosen
    If c = "Cash" Then
       MsgBox "Please deposit the money into the slot. only $5,$10,$20,$50 US currency will be accepted. Thankyou.", , "Deposit Money"
    End If
'displays the front page and hides the second page
 frmFrontPage.Show
 frmSecondPage.Hide
 
End Sub

Private Sub cmdFemaleSection_Click()
'hides this form and takes the user to the women form
frmSecondPage.Hide
frmWomen.Show

End Sub

Private Sub cmdMaleSection_Click()
'hides this form and takes the user to the Men form
frmSecondPage.Hide
frmMen.Show

End Sub


Private Sub cmdTrivia_Click()
'brings up the trivia form so the user can view it
    frmTrivia.Show
End Sub
