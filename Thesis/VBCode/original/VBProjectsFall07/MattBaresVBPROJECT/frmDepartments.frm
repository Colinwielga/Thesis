VERSION 5.00
Begin VB.Form frmDepartments 
   BackColor       =   &H80000001&
   Caption         =   "Target"
   ClientHeight    =   7305
   ClientLeft      =   4860
   ClientTop       =   4380
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   8220
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "Proceed to Checkout"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave Target"
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   6120
      Width           =   2895
   End
   Begin VB.CommandButton cmdElectronics 
      Caption         =   "Electronics"
      Height          =   735
      Left            =   5760
      TabIndex        =   4
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdToys 
      Caption         =   "Toys"
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdKitchen 
      Caption         =   "Kitchen"
      Height          =   735
      Left            =   5760
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdFurniture 
      Caption         =   "Furniture"
      Height          =   735
      Left            =   5760
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdClothing 
      Caption         =   "Clothing"
      Height          =   735
      Left            =   5760
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.PictureBox picBkgd 
      Height          =   7095
      Left            =   120
      Picture         =   "frmDepartments.frx":0000
      ScaleHeight     =   7035
      ScaleWidth      =   7875
      TabIndex        =   8
      Top             =   120
      Width           =   7935
      Begin VB.Label lblProceed 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Or, if you are finished shopping you may proceed to checkout."
         Height          =   615
         Left            =   960
         TabIndex        =   10
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label lblSelect 
         BackStyle       =   0  'Transparent
         Caption         =   "Please choose a department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   9
         Top             =   3120
         Width           =   3495
      End
      Begin VB.Shape shpEntrance 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   720
         Top             =   3000
         Width           =   3735
      End
   End
   Begin VB.Line Line8 
      X1              =   4440
      X2              =   2520
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line7 
      X1              =   5640
      X2              =   4440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line6 
      X1              =   5640
      X2              =   4440
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line5 
      X1              =   5640
      X2              =   4440
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line4 
      X1              =   4440
      X2              =   4440
      Y1              =   6480
      Y2              =   3600
   End
   Begin VB.Line Line3 
      X1              =   5640
      X2              =   4440
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      X1              =   4440
      X2              =   4440
      Y1              =   3600
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   4440
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblChoose 
      Caption         =   "Please choose a department."
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
End
Attribute VB_Name = "frmDepartments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheckOut_Click()
'This subroutine hides the department selection form and shows the checkout form.
'NOTE: this button is not available until an item has been added to the cart of any department.
frmDepartments.Hide
frmCheckout.Show
End Sub

Private Sub cmdClothing_Click()
'This subroutine hides the department selection form and shows the clothing department form.
frmDepartments.Hide
frmClothing.Show
End Sub

Private Sub cmdElectronics_Click()
'This subroutine hides the department selection form and shows the electronics department form.
frmDepartments.Hide
frmElectronics.Show
End Sub

Private Sub cmdFurniture_Click()
'This subroutine hides the department selection form and shows the furniture department form.
frmDepartments.Hide
frmFurniture.Show
End Sub

Private Sub cmdKitchen_Click()
'This subroutine hides the department selection form and shows the kithcen department form.
frmDepartments.Hide
frmKitchen.Show
End Sub

Private Sub cmdLeave_Click()
'This subroutine thanks the user for shopping via a msgbox, and quits the program.
MsgBox "Thanks for shopping at Target!  We hope you enjoyed our store.", , "Target"
End
End Sub

Private Sub cmdToys_Click()
'This subroutine hides the department selection form and shows the toys department form.
frmDepartments.Hide
frmToys.Show
End Sub

