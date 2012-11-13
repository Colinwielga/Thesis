VERSION 5.00
Begin VB.Form Selection 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton move 
      Caption         =   "Continue!"
      Height          =   375
      Left            =   2760
      TabIndex        =   29
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton tally 
      Caption         =   "Process totals"
      Height          =   375
      Left            =   960
      TabIndex        =   28
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox bclubskirt 
      Height          =   285
      Left            =   3720
      TabIndex        =   27
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox bclubfront 
      Height          =   285
      Left            =   3720
      TabIndex        =   26
      Text            =   "0"
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H00000040&
      Caption         =   "Option8"
      Height          =   255
      Left            =   3360
      TabIndex        =   25
      Top             =   4800
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000040&
      Caption         =   "Option1"
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox Eriskirt 
      Height          =   285
      Left            =   3720
      TabIndex        =   21
      Text            =   "0"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox eribump 
      Height          =   285
      Left            =   3720
      TabIndex        =   20
      Text            =   "0"
      Top             =   3000
      Width           =   495
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00000040&
      Caption         =   "Option7"
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   3360
      Width           =   255
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00000040&
      Caption         =   "Option6"
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox bombskirt 
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00000040&
      Caption         =   "Option5"
      Height          =   195
      Left            =   3360
      TabIndex        =   12
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00000040&
      Caption         =   "Option4"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox bombump 
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Text            =   "0"
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox bodykitpic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   4680
      ScaleHeight     =   4095
      ScaleWidth      =   6375
      TabIndex        =   4
      Top             =   840
      Width           =   6375
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000040&
      Caption         =   "Option3"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   4080
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000040&
      Caption         =   "Option2"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton showbomb 
      BackColor       =   &H00000040&
      Caption         =   "Option1"
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000040&
      Caption         =   "Buddy Club Side Skirts (2) - $349"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000040&
      Caption         =   "Buddy Club Front Bumper - $399"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   22
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000040&
      Caption         =   "Quantity"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000040&
      Caption         =   "View"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000040&
      Caption         =   "Erebuni Side Skirts (2) - $349"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000040&
      Caption         =   "Erebuni Front Bumper - $387"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000040&
      Caption         =   "Bomb Kit Side Skirts (2) - $249"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000040&
      Caption         =   "Bomb Kit front Bumper - $299"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000040&
      Caption         =   "Andy's Autosport Buddy Club Kit"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000040&
      Caption         =   "Erebuni Shogun Kit"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000040&
      Caption         =   "Andy's Autosport Bomb Kit"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      Caption         =   "Purchase a body kit for your ZX2."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : BodyKitOrder (Mike Seifert VB Project.vbp)
'Form Name : Selection (Form1.frm)
'Author: Mike Seifert
'Date Written: November 5, 2003
'Purpose of Form: To allow the user to see and select from various
        'body kit styles for purchase, and then to tally the total
        'cost of the purchase and compute shipping.

'message box containing instrutions.
Private Sub Form_Load()

Dim path As String
path = "App.path"

MsgBox "Enter the number of each part you intend to order.  Leave unused boxes with the value zero", , "Instructions"
End Sub

'enters the values from the text boxes into variables
Private Sub tally_Click()

'establishes a directory path for files used in this program


If bombump.Text > 0 Then
    aasbump = bombump.Text
Else
    aasbump = 0
End If

If bombskirt.Text > 0 Then
    aaskirt = bombskirt.Text
Else
    aaskirt = 0
End If

If eribump.Text > 0 Then
    shogubump = eribump.Text
Else
    shogubump = 0
End If

If Eriskirt.Text > 0 Then
    soskirt = Eriskirt.Text
Else
    soskirt = 0
End If

If bclubfront.Text > 0 Then
    budbump = bclubfront.Text
Else
    budbump = 0
End If

If bclubskirt.Text > 0 Then
    buddyskirt = bclubskirt.Text
Else
    buddyclubskirt = 0
End If

End Sub
'All of the option clicks display pictures of the parts to be purchased.
Private Sub Option1_Click()
bodykitpic.Picture = LoadPicture(path & "buddyclubfront.jpg")
End Sub

Private Sub Option2_Click()
bodykitpic.Picture = LoadPicture(path & "shogun.jpg")
End Sub

Private Sub Option3_Click()
bodykitpic.Picture = LoadPicture(path & "buddyclub.jpg")
End Sub

Private Sub Option4_Click()
bodykitpic.Picture = LoadPicture(path & "bombkit copy.jpg")
End Sub

Private Sub Option5_Click()
bodykitpic.Picture = LoadPicture(path & "aasskirts.jpg")
End Sub

Private Sub Option6_Click()
bodykitpic.Picture = LoadPicture(path & "shogunfront.jpg")
End Sub

Private Sub Option7_Click()
bodykitpic.Picture = LoadPicture(path & "shogunskirt.jpg")
End Sub

Private Sub Option8_Click()
bodykitpic.Picture = LoadPicture(path & "buddyclubskirt.jpg")
End Sub
Private Sub showbomb_Click()
bodykitpic.Picture = LoadPicture(path & "bombkit.jpg")
End Sub
'switches to the next form
Private Sub move_Click()
Checkout.Show
Selection.Hide
End Sub
