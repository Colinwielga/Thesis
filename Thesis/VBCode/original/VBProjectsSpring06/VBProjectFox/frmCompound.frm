VERSION 5.00
Begin VB.Form frmCompound 
   BackColor       =   &H00008000&
   Caption         =   "Pinnately or Palmately Compound Leaves"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAlt 
      Caption         =   "If your tree has pinnately compound, alternate leave: Click Here"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   10
      Top             =   7560
      Width           =   2535
   End
   Begin VB.CommandButton cmdAsh 
      Caption         =   "To find out about the Ash: Click Here"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   7680
      Width           =   2655
   End
   Begin VB.CommandButton cmdOhioBuckeye 
      Caption         =   "To find out about the Ohio Buckeye/ Horse Chesnut: Click Here"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdBox 
      Caption         =   "To find out about the Box elder/ Ash-leaved maple: Click Here"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdEndCompound 
      Caption         =   "End Progam"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturnfromCompound 
      Caption         =   "Return to Beginning of Program"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Image imgAsh2 
      Height          =   5250
      Left            =   0
      Picture         =   "frmCompound.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image imgBox2 
      Height          =   3915
      Left            =   480
      Picture         =   "frmCompound.frx":4FA32
      Top             =   0
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Image imgOhio2 
      Height          =   3855
      Left            =   360
      Picture         =   "frmCompound.frx":741A0
      Top             =   120
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Has pinnately compound and alternate  leaves "
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   5760
      TabIndex        =   7
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Image imgAlternate 
      Height          =   3255
      Left            =   6120
      Picture         =   "frmCompound.frx":9801E
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   4695
   End
   Begin VB.Label lblBox 
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmCompound.frx":F14A8
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   5280
      TabIndex        =   5
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label lblAsh 
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pinnately compound and opposite leaves with 5-11 leaflets (although sometimes may have less) and short petioles (Ashes)"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   1920
      TabIndex        =   4
      Top             =   5400
      Width           =   3855
   End
   Begin VB.Label lblPalmate 
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Palmately Compound Leaves (Ohio Buckeye or Horse Chesnut)"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label lblTitleCompound 
      BackStyle       =   0  'Transparent
      Caption         =   "Are leaves palmately or pinnately compound?"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   3120
      TabIndex        =   2
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Image imgAsh 
      Height          =   3120
      Left            =   0
      Picture         =   "frmCompound.frx":F1545
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   4440
   End
   Begin VB.Image imgBox 
      Height          =   3540
      Left            =   6360
      Picture         =   "frmCompound.frx":331587
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4680
   End
   Begin VB.Image imgOhioBuckeye 
      Height          =   3720
      Left            =   0
      Picture         =   "frmCompound.frx":44E209
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4560
   End
End
Attribute VB_Name = "frmCompound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmCompound(frmCompound.frm)
'Author: Kelly Fox
'Date Written:3/20/2006
'This form allow the use to identify trees with compound leaves

Private Sub cmdAlt_Click()
    'Go to the form frmPinnatelycompound
    frmCompound.Hide
    frmPinnatelyCompound.Show
End Sub

Private Sub cmdAsh_Click()
    'this long list of visible and invisible objects allows for the user to better recognize the important information by temporaily taking out most of the background objects
    imgAsh2.Visible = True
    imgOhioBuckeye.Visible = False
    cmdAsh.Visible = False
    imgAlternate.Visible = False
    imgBox.Visible = False
    imgAsh.Visible = False
    cmdOhioBuckeye.Visible = False
    lblAsh.Visible = False
    lblAlt.Visible = False
    lblPalmate.Visible = False
    lblTitleCompound.Visible = False
    lblBox.Visible = False
    cmdAlt.Visible = False
    cmdBox.Visible = False
    cmdEndCompound.Visible = False
    cmdReturnfromCompound.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Fraxinus, and is commonly known as the Ash", , "Genus: Fraxinus"
    imgAsh2.Visible = False
    imgOhioBuckeye.Visible = True
    cmdAsh.Visible = True
    imgAlternate.Visible = True
    imgBox.Visible = True
    imgAsh.Visible = True
    cmdOhioBuckeye.Visible = True
    cmdAlt.Visible = True
    cmdBox.Visible = True
    cmdEndCompound.Visible = True
    lblAsh.Visible = True
    lblAlt.Visible = True
    lblPalmate.Visible = True
    lblTitleCompound.Visible = True
    lblBox.Visible = True
    cmdReturnfromCompound.Visible = True
End Sub

Private Sub cmdBox_Click()
    'This control once again uses a long list of visibles and invisibles to make it easier to interpret the data
    imgBox2.Visible = True
    imgOhioBuckeye.Visible = False
    cmdAsh.Visible = False
    imgAlternate.Visible = False
    imgBox.Visible = False
    imgAsh.Visible = False
    cmdOhioBuckeye.Visible = False
    lblAsh.Visible = False
    lblAlt.Visible = False
    lblPalmate.Visible = False
    lblTitleCompound.Visible = False
    lblBox.Visible = False
    cmdAlt.Visible = False
    cmdBox.Visible = False
    cmdEndCompound.Visible = False
    cmdReturnfromCompound.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Acer, and is commonly known as the Box elder or the Ash-leaved Maple ", , "Genus: Acer"
    imgBox2.Visible = False
    imgOhioBuckeye.Visible = True
    cmdAsh.Visible = True
    imgAlternate.Visible = True
    imgBox.Visible = True
    imgAsh.Visible = True
    cmdOhioBuckeye.Visible = True
    cmdAlt.Visible = True
    cmdBox.Visible = True
    cmdEndCompound.Visible = True
    lblAsh.Visible = True
    lblAlt.Visible = True
    lblPalmate.Visible = True
    lblTitleCompound.Visible = True
    lblBox.Visible = True
    cmdReturnfromCompound.Visible = True
End Sub

Private Sub cmdEndCompound_Click()
    'Ends program
    End
End Sub

Private Sub cmdOhioBuckeye_Click()
    imgOhio2.Visible = True
    imgOhioBuckeye.Visible = False
    cmdAsh.Visible = False
    imgAlternate.Visible = False
    imgBox.Visible = False
    imgAsh.Visible = False
    cmdOhioBuckeye.Visible = False
    lblAsh.Visible = False
    lblAlt.Visible = False
    lblPalmate.Visible = False
    lblTitleCompound.Visible = False
    lblBox.Visible = False
    cmdAlt.Visible = False
    cmdBox.Visible = False
    cmdEndCompound.Visible = False
    cmdReturnfromCompound.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Aesculus, and is commonly known as the Ohio Buckeye or the Horse Chesnut ", , "Genus: Aesculus"
    imgOhio2.Visible = False
    imgOhioBuckeye.Visible = True
    cmdAsh.Visible = True
    imgAlternate.Visible = True
    imgBox.Visible = True
    imgAsh.Visible = True
    cmdOhioBuckeye.Visible = True
    cmdAlt.Visible = True
    cmdBox.Visible = True
    cmdEndCompound.Visible = True
    lblAsh.Visible = True
    lblAlt.Visible = True
    lblPalmate.Visible = True
    lblTitleCompound.Visible = True
    lblBox.Visible = True
    cmdReturnfromCompound.Visible = True
End Sub

Private Sub cmdReturnfromCompound_Click()
    frmCompound.Hide
    frmMinnesotaTrees.Show
End Sub

