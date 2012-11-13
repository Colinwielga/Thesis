VERSION 5.00
Begin VB.Form frmJeans 
   BackColor       =   &H00808000&
   Caption         =   "Jeans"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheckout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Proceed to Checkout"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdHandbags 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Handbags"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdShoes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shoes"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdDresses 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dresses"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtChoose 
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   6720
      Width           =   735
   End
   Begin VB.PictureBox picBootcutJean 
      Height          =   2655
      Left            =   1800
      Picture         =   "frmJeans.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.PictureBox picSkinnyJean 
      Height          =   2655
      Left            =   1800
      Picture         =   "frmJeans.frx":11DBA
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblChoose 
      BackColor       =   &H00808000&
      Caption         =   "Would You Like #1 or #2?"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   6720
      Width           =   3615
   End
   Begin VB.Label lblJeans 
      BackColor       =   &H00808000&
      Caption         =   "J E A N S"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblBootcutJean 
      BackColor       =   &H00808000&
      Caption         =   "#2 - Bootcut Jean"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label lblSkinnyJean 
      BackColor       =   &H00808000&
      Caption         =   "#1 - Skinny Jean"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmJeans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form will allow the user to choose between two types of jeans
'and add the desired pair to his or her shopping cart.  Then the user
'will receive an updated balance on the subtotal of his or her purchase.


'Move from current form to checkout form using visible variable
Private Sub cmdCheckout_Click()
    frmJeans.Visible = False
    frmCheckout.Visible = True
End Sub
'Move from current form to dresses form using visible variable
Private Sub cmdDresses_Click()
    frmJeans.Visible = False
    frmDresses.Visible = True
End Sub
'Move from current form to handbags form using visible variable
Private Sub cmdHandbags_Click()
    frmJeans.Visible = False
    frmHandbags.Visible = True
End Sub
'Move form current form to shoes form using visible variable
Private Sub cmdShoes_Click()
    frmJeans.Visible = False
    frmShoes.Visible = True
End Sub
'The user may choose between the jeans by inputing his of her choice in
'the input box.  Then a messagebox will inform the user of his or her
'current subtotal
Private Sub txtChoose_Change()
Dim Choose As Integer
Choose = txtChoose.Text


    If Choose = 1 Then
        ShoppingCart = ShoppingCart + 163
        Items = Items + 1
        MsgBox "Your subtotal thus far is  " + FormatCurrency(ShoppingCart), , "Subtotal"
    ElseIf Choose = 2 Then
        ShoppingCart = ShoppingCart + 175
        Items = Items + 1
        MsgBox "Your subtotal thus far is  " + FormatCurrency(ShoppingCart), , "Subtotal"
    End If
End Sub
