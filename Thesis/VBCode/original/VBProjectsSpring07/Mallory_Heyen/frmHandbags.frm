VERSION 5.00
Begin VB.Form frmHandbags 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Handbags"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheckout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Proceed to Checkout"
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdJeans 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jeans"
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdShoes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shoes"
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDresses 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dresses"
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtChoose 
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.PictureBox picHobo 
      Height          =   2655
      Left            =   3720
      Picture         =   "frmHandbags.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.PictureBox picClutch 
      Height          =   2655
      Left            =   840
      Picture         =   "frmHandbags.frx":11962
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblChoose 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Would You Like #1 or #2?"
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label lblHandbags 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Handbags"
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label lblHobo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "#2 H O B O"
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5880
      TabIndex        =   3
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblClutch 
      BackColor       =   &H00FFC0C0&
      Caption         =   "#1 C L U T C H"
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   255
   End
End
Attribute VB_Name = "frmHandbags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The user can choose from two items and monitor his or her subtotal
'while shopping

'Move from current form to checkout form using visible variable
Private Sub cmdCheckout_Click()
    frmHandbags.Visible = False
    frmCheckout.Visible = True
End Sub
'Move from current form to dresses form using visible variable
Private Sub cmdDresses_Click()
    frmHandbags.Visible = False
    frmDresses.Visible = True
End Sub
'Move from current forn to jeans form using visible variable
Private Sub cmdJeans_Click()
    frmHandbags.Visible = False
    frmJeans.Visible = True
End Sub
'Move from current form to shoes form using visible variable
Private Sub cmdShoes_Click()
    frmHandbags.Visible = False
    frmShoes.Visible = True
End Sub
'The user is able to indicate which item he or she would like to
'purchse.  Then a messagebox will indicate what the current
'subtotal is.
Private Sub txtChoose_Change()
Dim Choose As Integer
Choose = txtChoose.Text


    If Choose = 1 Then
        ShoppingCart = ShoppingCart + 215
        Items = Items + 1
        MsgBox "Your subtotal thus far is  " + FormatCurrency(ShoppingCart), , "Subtotal"
    ElseIf Choose = 2 Then
        ShoppingCart = ShoppingCart + 438
        Items = Items + 1
        MsgBox "Your subtotal thus far is  " + FormatCurrency(ShoppingCart), , "Subtotal"
    End If
End Sub


