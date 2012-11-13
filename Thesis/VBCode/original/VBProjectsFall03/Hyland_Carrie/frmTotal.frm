VERSION 5.00
Begin VB.Form frmTotal 
   BackColor       =   &H00C000C0&
   Caption         =   "Total"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCost 
      Caption         =   "Click to see your Total Price!"
      Height          =   855
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Click to Buy!"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdReOrder1 
      Caption         =   "Want another?  Click here to order!"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      Height          =   735
      Left            =   3600
      TabIndex        =   1
      Top             =   2880
      Width           =   1815
   End
   Begin VB.PictureBox picTotal 
      Height          =   2415
      Left            =   2640
      ScaleHeight     =   2355
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuy_Click()
MsgBox ("Thank you for ordering, we appreciate your business!  Enjoy!")
End Sub

Private Sub cmdCost_Click()
If Burrito = A Then
        picTotal.Print A;
        picTotal.Print F
    ElseIf Burrito = B Then
        picTotal.Print B;
        picTotal.Print G
End If
    
    
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReOrder1_Click()
frmTotal.Hide
frmOrder.Show
End Sub
