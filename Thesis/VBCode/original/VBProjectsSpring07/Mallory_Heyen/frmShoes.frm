VERSION 5.00
Begin VB.Form frmShoes 
   BackColor       =   &H00C000C0&
   Caption         =   "Shoes"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheckout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Proceed to Checkout"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdJeans 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jeans"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdHandbags 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Handbags"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdDresses 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dresses"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtChoose 
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.PictureBox picPumps 
      Height          =   2655
      Left            =   4800
      Picture         =   "frmShoes.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.PictureBox picWedges 
      Height          =   2655
      Left            =   1800
      Picture         =   "frmShoes.frx":11962
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblShoes 
      BackColor       =   &H0000FFFF&
      Caption         =   "SHOES"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblChoose 
      BackColor       =   &H000080FF&
      Caption         =   "Would you like #1 or #2?"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      TabIndex        =   5
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label lblPumps 
      BackColor       =   &H00C0C000&
      Caption         =   "#2 - Pumps"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblWedges 
      BackColor       =   &H00C0C000&
      Caption         =   "#1 - Wedges"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   4080
      Width           =   2055
   End
End
Attribute VB_Name = "frmShoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program will allow the user to choose between two pairs of
'shoes and then update the user on the subtotal of the purchase

'Move from current form to chechout form using visible variable
Private Sub cmdCheckout_Click()
    frmShoes.Visible = False
    frmCheckout.Visible = True
End Sub
'Move from current form to dresses form using visible variable
Private Sub cmdDresses_Click()
    frmShoes.Visible = False
    frmDresses.Visible = True
End Sub
'Move from current form to handbags form using visible variable
Private Sub cmdHandbags_Click()
    frmShoes.Visible = False
    frmHandbags.Visible = True
End Sub
'Move from current form to jeans form using visible variable
Private Sub cmdJeans_Click()
    frmShoes.Visible = False
    frmJeans.Visible = True
End Sub


'The user indicates which item he or she would like in the text box
'and then is updated on the subtotal of his or her purchases
Private Sub txtChoose_Change()
Dim Choose As Integer
Choose = txtChoose.Text


    If Choose = 1 Then
        ShoppingCart = ShoppingCart + 225
        Items = Items + 1
        MsgBox "Your subtotal thus far is  " + FormatCurrency(ShoppingCart), , "Subtotal"
    ElseIf Choose = 2 Then
        ShoppingCart = ShoppingCart + 275
        Items = Items + 1
        MsgBox "Your subtotal thus far is  " + FormatCurrency(ShoppingCart), , "Subtotal"
    End If
End Sub
