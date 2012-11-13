VERSION 5.00
Begin VB.Form frm_roller 
   BackColor       =   &H0080FFFF&
   Caption         =   "Roll Simulator by Sean Wasmund"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_bacc 
      Caption         =   "Baccarat"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmd_opt 
      Caption         =   "Options"
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmd_roll 
      Caption         =   "Roll"
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmd_quit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmd_reset 
      Caption         =   "Clear"
      Height          =   735
      Left            =   3960
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox results 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   360
      ScaleHeight     =   1995
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   1200
      Width           =   5895
   End
End
Attribute VB_Name = "frm_roller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_bacc_Click()        'opens baccarat
frm_roller.Hide
frm_bacc.Show
End Sub

Private Sub cmd_opt_Click()         'opens options menu
results.Cls
cmd_roll.Enabled = True
frm_roller.Enabled = False
frm_options.Show
End Sub

Private Sub cmd_quit_Click()        'quits program
End
End Sub
Private Sub cmd_reset_Click()       'clears output
results.Cls
End Sub
Private Sub cmd_roll_Click()        'rolls selected dice
Dim dicetotal As Integer
Dim i As Integer
Dim rollres() As Integer
ReDim rollres(1 To numdie)

For i = 1 To numdie
    rollres(i) = Int(high * Rnd) + 1        'generates random numbers
    results.Print rollres(i),
    dicetotal = dicetotal + rollres(i)
Next i
results.Print ""
results.Print Tab(69); "Total: "; dicetotal 'totals dice

End Sub

Private Sub Form_Load()
Randomize
cmd_roll.Enabled = False
End Sub
