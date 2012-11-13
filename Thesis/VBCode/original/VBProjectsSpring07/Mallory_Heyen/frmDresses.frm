VERSION 5.00
Begin VB.Form frmDresses 
   BackColor       =   &H00000000&
   Caption         =   "Dresses"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheckout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Proceed to Checkout"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton cmdShoes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shoes"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdJeans 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jeans"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdHandbags 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Handbags"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtChoose 
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   5640
      Width           =   735
   End
   Begin VB.PictureBox picCocktail 
      Height          =   2655
      Left            =   2040
      Picture         =   "frmDresses.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.PictureBox picGown 
      Height          =   2655
      Left            =   4200
      Picture         =   "frmDresses.frx":11962
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblChoose 
      BackColor       =   &H00000000&
      Caption         =   "Would You Like #1 or #2?"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   5760
      Width           =   4455
   End
   Begin VB.Label lblDresses 
      BackColor       =   &H00000000&
      Caption         =   "Dresses"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblGown 
      BackColor       =   &H00000000&
      Caption         =   "#2 - Gown"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblCocktail 
      BackColor       =   &H00000000&
      Caption         =   "#1 - Cocktail"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "frmDresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form will allow the user to choose between two dress option and then
'add the price of the chosen dress to the shopping cart.  If the user does not
'like either dress he or she can return to the front page and continue shopping.


'Move from this form to the checkout form using the visible variable
Private Sub cmdCheckout_Click()
    frmDresses.Visible = False
    frmCheckout.Visible = True
End Sub

'Move from this form to the Handbags form using the visible variable
Private Sub cmdHandbags_Click()
    frmDresses.Visible = False
    frmHandbags.Visible = True
    
End Sub


'Move from this form to the jeans form using the visible variable
Private Sub cmdJeans_Click()
    frmDresses.Visible = False
    frmJeans.Visible = True
    
End Sub
'Move from this form to the shoes form using the visible variable
Private Sub cmdShoes_Click()
    frmDresses.Visible = False
    frmShoes.Visible = True
    
End Sub

'The user will enter which item they would like to purchase in the
'input box, which will then show the user his or her subtotal thus
'far
Private Sub txtChoose_Change()
Dim Choose As Integer
Choose = txtChoose.Text


    If Choose = 1 Then
        ShoppingCart = ShoppingCart + 330
        Items = Items + 1
        MsgBox "Your subtotal thus far is  " + FormatCurrency(ShoppingCart), , "Subtotal"
    ElseIf Choose = 2 Then
        ShoppingCart = ShoppingCart + 410
        Items = Items + 1
        MsgBox "Your subtotal thus far is  " + FormatCurrency(ShoppingCart), , "Subtotal"
    End If
    
End Sub
