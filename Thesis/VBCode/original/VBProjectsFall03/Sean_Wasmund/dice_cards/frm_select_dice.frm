VERSION 5.00
Begin VB.Form frm_options 
   BackColor       =   &H0080C0FF&
   Caption         =   "Select Dice by Sean Wasmund"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_sides3 
      Caption         =   "8 Sides"
      Height          =   735
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmd_done 
      Caption         =   "Done"
      Height          =   2655
      Left            =   3360
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmd_reset1 
      Caption         =   "Reset"
      Height          =   1695
      Left            =   3360
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmd_r5 
      Caption         =   "Roll 5 Dice"
      Height          =   735
      Left            =   1920
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmd_r4 
      Caption         =   "Roll 4 Dice"
      Height          =   735
      Left            =   1920
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmd_r3 
      Caption         =   "Roll 3 Dice"
      Height          =   735
      Left            =   1920
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmd_r2 
      Caption         =   "Roll 2 Dice"
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmd_r1 
      Caption         =   "Roll 1 Die"
      Height          =   735
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmd_sides6 
      Caption         =   "20 Sides"
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmd_sides5 
      Caption         =   "12 Sides"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmd_sides4 
      Caption         =   "10 Sides"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmd_sides2 
      Caption         =   "6 Sides"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmd_sides1 
      Caption         =   "4 Sides"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frm_options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_done_Click()        'returns to roller form
cmd_sides1.Enabled = True
cmd_sides2.Enabled = True
cmd_sides3.Enabled = True
cmd_sides4.Enabled = True
cmd_sides5.Enabled = True
cmd_sides6.Enabled = True

cmd_r1.Enabled = True
cmd_r2.Enabled = True
cmd_r3.Enabled = True
cmd_r4.Enabled = True
cmd_r5.Enabled = True

frm_roller.Enabled = True
frm_options.Hide

End Sub


Private Sub cmd_r1_Click()      'rolls 1 die
numdie = 1
cmd_r1.Enabled = False
cmd_r2.Enabled = False
cmd_r3.Enabled = False
cmd_r4.Enabled = False
cmd_r5.Enabled = False
End Sub

Private Sub cmd_r2_Click()      'rolls 2 dice
numdie = 2
cmd_r1.Enabled = False
cmd_r2.Enabled = False
cmd_r3.Enabled = False
cmd_r4.Enabled = False
cmd_r5.Enabled = False
End Sub

Private Sub cmd_r3_Click()      'rolls 3 dice
numdie = 3
cmd_r1.Enabled = False
cmd_r2.Enabled = False
cmd_r3.Enabled = False
cmd_r4.Enabled = False
cmd_r5.Enabled = False
End Sub

Private Sub cmd_r4_Click()      'rolls 4 dice
numdie = 4
cmd_r1.Enabled = False
cmd_r2.Enabled = False
cmd_r3.Enabled = False
cmd_r4.Enabled = False
cmd_r5.Enabled = False
End Sub

Private Sub cmd_r5_Click()      'rolls 5 dice
numdie = 5
cmd_r1.Enabled = False
cmd_r2.Enabled = False
cmd_r3.Enabled = False
cmd_r4.Enabled = False
cmd_r5.Enabled = False
End Sub

Private Sub cmd_reset1_Click()  'resets options form
cmd_sides1.Enabled = True
cmd_sides2.Enabled = True
cmd_sides3.Enabled = True
cmd_sides4.Enabled = True
cmd_sides5.Enabled = True
cmd_sides6.Enabled = True

cmd_r1.Enabled = True
cmd_r2.Enabled = True
cmd_r3.Enabled = True
cmd_r4.Enabled = True
cmd_r5.Enabled = True
End Sub

Private Sub cmd_sides1_Click()  'selects 4 sided die
high = 4
cmd_sides1.Enabled = False
cmd_sides2.Enabled = False
cmd_sides3.Enabled = False
cmd_sides4.Enabled = False
cmd_sides5.Enabled = False
cmd_sides6.Enabled = False
End Sub

Private Sub cmd_sides2_Click()  'selects 6 sided die
high = 6
cmd_sides1.Enabled = False
cmd_sides2.Enabled = False
cmd_sides3.Enabled = False
cmd_sides4.Enabled = False
cmd_sides5.Enabled = False
cmd_sides6.Enabled = False
End Sub

Private Sub cmd_sides3_Click()  'selects 8 sided die
high = 8
cmd_sides1.Enabled = False
cmd_sides2.Enabled = False
cmd_sides3.Enabled = False
cmd_sides4.Enabled = False
cmd_sides5.Enabled = False
cmd_sides6.Enabled = False
End Sub

Private Sub cmd_sides4_Click()      'selects 10 sided die
high = 10
cmd_sides1.Enabled = False
cmd_sides2.Enabled = False
cmd_sides3.Enabled = False
cmd_sides4.Enabled = False
cmd_sides5.Enabled = False
cmd_sides6.Enabled = False
End Sub

Private Sub cmd_sides5_Click()      'selects 12 sided die
high = 12
cmd_sides1.Enabled = False
cmd_sides2.Enabled = False
cmd_sides3.Enabled = False
cmd_sides4.Enabled = False
cmd_sides5.Enabled = False
cmd_sides6.Enabled = False
End Sub

Private Sub cmd_sides6_Click()     'selects 20 sided die
high = 20
cmd_sides1.Enabled = False
cmd_sides2.Enabled = False
cmd_sides3.Enabled = False
cmd_sides4.Enabled = False
cmd_sides5.Enabled = False
cmd_sides6.Enabled = False
End Sub

Private Sub Form_Load()

End Sub
