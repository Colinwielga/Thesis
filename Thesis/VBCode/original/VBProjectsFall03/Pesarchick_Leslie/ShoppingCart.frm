VERSION 5.00
Begin VB.Form frmShoppingCart 
   BackColor       =   &H00800080&
   Caption         =   "Shopping Cart"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdBacktoMain 
      Caption         =   "Back to Main Meun"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   6600
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF80FF&
      Height          =   7335
      Left            =   1800
      ScaleHeight     =   7275
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmShoppingCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBacktoMain_Click()
    frmShoesetc.Show
    frmShoppingCart.Hide
End Sub
