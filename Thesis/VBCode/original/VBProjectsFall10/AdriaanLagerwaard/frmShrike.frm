VERSION 5.00
Begin VB.Form frmShrike 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tyranid Shrike Brood"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19500
   LinkTopic       =   "Form1"
   ScaleHeight     =   12210
   ScaleWidth      =   19500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShrikeToal 
      Caption         =   "Total Points Spent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      TabIndex        =   19
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdShrikeTotalCls 
      Caption         =   "Clear Points Spent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   18
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheckBoxCls 
      Caption         =   "Clear Check Boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      TabIndex        =   17
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdShrikeCls 
      Caption         =   "Clear Number of Tyranid Shrikes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CheckBox Check8 
      Caption         =   "An additional Set of Scything Talons Free"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CheckBox Check7 
      Caption         =   "BoneSword and Lash Whips 15pts/each"
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CheckBox Check6 
      Caption         =   "A Pair of BoneSwords 10pts/each"
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Deathspittes 5pts/each"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Spinfists Free"
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Rending Claws 5pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   3615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Venom Cannon 15pts"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   3615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Barbed Strangler 10pts"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3615
   End
   Begin VB.PictureBox picShrike 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   5835
      TabIndex        =   4
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox txtShrike 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdShrikeBack 
      Caption         =   "Back To Fast Attack"
      Height          =   975
      Left            =   5400
      TabIndex        =   0
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Equiped With:              -Devourer                      -Reinforced Chitin        -Scything Talons           -Wings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7680
      TabIndex        =   20
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "The entire Brood may replace its Devourers with:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   10
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "The entire Brood may replace its Scything Talons for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "One Tyranid Shike in the Brood may replace its Devourer with:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "3-9 per Brood:  35 pts/Each"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Tyranid Shrike Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmShrike"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
EquipShrike = EquipShrike + BarbedStrangler

End Sub

Private Sub Check2_Click()
EquipShrike = EquipShrike + VCannon

End Sub

Private Sub Check3_Click()
NumberShrike = txtShrike
EquipShrike = EquipShrike + (RendingClaws * NumberShrike)

End Sub

Private Sub Check5_Click()
NumberShrike = txtShrike
EquipShrike = EquipShrike + (Deathspitter * NumberShrike)
End Sub

Private Sub Check6_Click()
NumberShrike = txtShrike
EquipShrike = EquipShrike + (BoneSword * NumberShrike)

End Sub

Private Sub Check7_Click()
NumberShrike = txtShrike
EquipShrike = EquipShrike + (ImplantAttack * NumberShrike)

End Sub

Private Sub cmdCheckBoxCls_Click()

    If Check1 = 1 Then
        EquipShrike = EquipShrike - BarbedStrangler
    End If
Check1 = 0
    If Check2 = 1 Then
        EquipShrike = EquipShrike - VCannon
    End If
Check2 = 0
Check3 = 0
Check4 = 0
Check5 = 0
Check6 = 0
Check7 = 0
Check8 = 0

End Sub

Private Sub cmdShrikeBack_Click()
frmShrike.Hide
frmFastAttack.Show

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Shrike" Then
            picShrike.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picShrike.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picShrike.Print "Sorry."
    End If

    
End Sub

Private Sub cmdShrikeCls_Click()
txtShrike = 0

End Sub

Private Sub cmdShrikeToal_Click()
NumberShrike = txtShrike
ShrikeTotal = ShrikeTotal + (35 * NumberShrike)

MsgBox "Your Total Points Spent on Tyranid Shrikes is " & (ShrikeTotal + EquipShrike)

End Sub

Private Sub cmdShrikeTotalCls_Click()
ShrikeTotal = 0
EquipShrike = 0
txtShrike = 0

End Sub
