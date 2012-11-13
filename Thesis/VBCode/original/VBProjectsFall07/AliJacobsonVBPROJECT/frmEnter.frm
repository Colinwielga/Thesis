VERSION 5.00
Begin VB.Form frmEnter 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWelcome 
      BackColor       =   &H0000FFFF&
      Caption         =   "Welcome to Cub Foods Shopping.  Please Click Here for Instructions to begin shopping."
      Height          =   2415
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton cmdCheckOut 
      BackColor       =   &H00FF00FF&
      Caption         =   "Proceed To Check Out"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   5175
   End
   Begin VB.CommandButton cmdShopFrozen 
      BackColor       =   &H00FFFF80&
      Caption         =   "Shop Frozen Foods"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   5175
   End
   Begin VB.CommandButton cmdShopBakery 
      BackColor       =   &H000080FF&
      Caption         =   "Shop Bakery"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   5175
   End
   Begin VB.CommandButton cmdShopProduce 
      BackColor       =   &H0000FF00&
      Caption         =   "Shop Produce"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   6120
      Picture         =   "frmEnter.frx":0000
      Top             =   720
      Width           =   3465
   End
End
Attribute VB_Name = "frmEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnter_Click()
'this button takes the user to the produce form

frmProduce.Show
frmEnter.Hide
frmBakery.Hide
frmFrozen.Hide
frmCheckOut.Hide

End Sub

Private Sub cmdCheckOut_Click()
'this button takes the user to the check out form

frmProduce.Hide
frmBakery.Hide
frmFrozen.Hide
    
    If RunningTotal > 0 Then 'this makes it so the user can not check out if they haven't added any items to their cart
        frmCheckOut.Show
    Else
        MsgBox "You have not purchased anything, please continue shopping"
        frmEnter.Show
    End If

    
    

End Sub

Private Sub cmdShopBakery_Click()
'this button takes the user to the bakery form

frmBakery.Show
frmEnter.Hide
frmProduce.Hide
frmFrozen.Hide
frmCheckOut.Hide

End Sub

Private Sub cmdShopFrozen_Click()
'This button takes the user to the Frozen food form

frmEnter.Hide
frmProduce.Hide
frmBakery.Hide
frmFrozen.Show
frmCheckOut.Hide

End Sub

Private Sub cmdShopProduce_Click()
'this button takes the user to the Produce form

frmProduce.Show
frmBakery.Hide
frmFrozen.Hide
frmCheckOut.Hide
frmBakery.Hide

End Sub

Private Sub cmdWelcome_Click()
'this button asks for the user's name and explains the purpose of the program and also allows for the user to acces the other buttons

CustomerName = InputBox("Please Enter Your Name")

MsgBox "Hello " & CustomerName & " and thank you for shopping at Cub Foods, the purpose of this is program is to make shopping and calculating your totals easier than in the store and allow you to make your purchases through this program."

cmdShopProduce.Enabled = True
cmdShopFrozen.Enabled = True
cmdShopBakery.Enabled = True
cmdCheckOut.Enabled = True




End Sub
