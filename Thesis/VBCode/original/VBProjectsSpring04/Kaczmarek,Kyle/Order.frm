VERSION 5.00
Begin VB.Form frmorder 
   BackColor       =   &H00000000&
   Caption         =   "Order"
   ClientHeight    =   7710
   ClientLeft      =   1320
   ClientTop       =   1305
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11265
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   120
      Picture         =   "Order.frx":0000
      ScaleHeight     =   6315
      ScaleWidth      =   10755
      TabIndex        =   1
      Top             =   120
      Width           =   10815
   End
   Begin VB.CommandButton cmdjersey 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Order Jersey"
      Height          =   855
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   2415
   End
End
Attribute VB_Name = "frmorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : BaseballUniforms (BaseballUniforms.vbp)
'Form Name : Order (Order.frm)
'Author: Kyle Kaczmarek
'Date Written: March 6, 2004
'Purpose of Form:'This allows the user to access the other forms
                 'which enables them to place orders



Private Sub cmdjersey_Click()
frmcleats.Hide 'closes the cleats form
frmjerseys.Show 'shows the jerseys form
frmhats.Hide 'closes the hat form
frmorder.Hide 'closes the order form
frmpants.Hide 'closes the pants form
frmfinal.Hide 'closes the final form
End Sub

