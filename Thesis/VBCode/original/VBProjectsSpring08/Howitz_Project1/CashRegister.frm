VERSION 5.00
Begin VB.Form frmTill 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   FillColor       =   &H00FFFF80&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdpay 
      Caption         =   "PAY UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmdloginbut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login Please"
      Height          =   975
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.CommandButton cmdSoup 
      Caption         =   "Soup"
      Height          =   1575
      Left            =   3240
      TabIndex        =   7
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton cmdBakery 
      Caption         =   "Bakery"
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton cmdSnack 
      Caption         =   "Snack"
      Height          =   1575
      Left            =   5880
      TabIndex        =   5
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton cmdDeli 
      Caption         =   "Deli"
      Height          =   1575
      Left            =   5880
      TabIndex        =   4
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdSalad 
      Caption         =   "Salad"
      Height          =   1575
      Left            =   5880
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdDrink 
      Caption         =   "Beverages"
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdPizza 
      Caption         =   "Pizza"
      Height          =   1575
      Left            =   3240
      TabIndex        =   1
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdGrill 
      Caption         =   "Grill"
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Brought to you by..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   10
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   6360
      Picture         =   "CashRegister.frx":0000
      Top             =   240
      Width           =   3600
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H80000009&
      Caption         =   "SEXTON DINING"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmTill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sexton Cash Register
'Form Name:  frmTill
'Louis Howitz
'March 31, 2008
'This is the form that contains all the food categories available
'at Sexton Dining.  Each button will lead to a different menu of
'food to purchase.

Private Sub cdmBakery_Click()
    frmTill.Hide
    frmBakery.Show
End Sub

Private Sub cmdBakery_Click()
    frmTill.Hide
    frmBakery.Show
    
End Sub

Private Sub cmdDeli_Click()
    frmTill.Hide
    frmDeli.Show
    
End Sub

Private Sub cmdDrink_Click()
    frmTill.Hide
    frmBev.Show
    
End Sub

Private Sub cmdGrill_Click()

    frmTill.Hide
    frmGrill.Show
    
End Sub

Private Sub cmdloginbut_Click()
    
    frmTill.Hide
    frmEnter.Show
    
End Sub

Private Sub cmdPay_Click()
    
    frmTill.Hide
    frmPay.Show
    
End Sub

Private Sub cmdPizza_Click()
    
    frmTill.Hide
    frmPizza.Show
    
End Sub

Private Sub cmdSalad_Click()
    frmTill.Hide
    frmsalad.Show
End Sub

Private Sub cmdSnack_Click()
    frmTill.Hide
    frmSnack.Show
    
End Sub

Private Sub cmdSoup_Click()
    frmTill.Hide
    frmSoup.Show
End Sub

Private Sub Form_Load()
    Items = 0
    MsgBox "Please login before you begin", , "Login"
    
End Sub
