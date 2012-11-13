VERSION 5.00
Begin VB.Form frmWarrior 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tyranid Warrior Brood"
   ClientHeight    =   11505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18465
   LinkTopic       =   "Form6"
   ScaleHeight     =   11505
   ScaleWidth      =   18465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   5055
      Left            =   9960
      ScaleHeight     =   4995
      ScaleWidth      =   5595
      TabIndex        =   24
      Top             =   600
      Width           =   5655
   End
   Begin VB.CommandButton cmdWarriorTotal 
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
      Left            =   7200
      TabIndex        =   23
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdWarriorTotalCls 
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
      Left            =   4800
      TabIndex        =   22
      Top             =   6000
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
      Left            =   2520
      TabIndex        =   21
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdNumberWarriorCls 
      Caption         =   "Clear Number of Warriors"
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
      Left            =   240
      TabIndex        =   20
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Toxin Sacs 5pts/each"
      Height          =   615
      Left            =   4200
      TabIndex        =   18
      Top             =   5280
      Width           =   3735
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Adrenal Glands 5pts/each"
      Height          =   735
      Left            =   4200
      TabIndex        =   17
      Top             =   4800
      Width           =   3735
   End
   Begin VB.CheckBox Check9 
      Caption         =   "additional set of Scything Talons Free"
      Height          =   615
      Left            =   4200
      TabIndex        =   16
      Top             =   3960
      Width           =   3735
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Deathspitter 5pts/each"
      Height          =   615
      Left            =   4200
      TabIndex        =   15
      Top             =   3480
      Width           =   3735
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Sinefists Free"
      Height          =   615
      Left            =   4200
      TabIndex        =   14
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Rending Claws Free"
      Height          =   615
      Left            =   4200
      TabIndex        =   13
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Venom Cannon 15pts"
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   3975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Barbed Strangler 10pts"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   4680
      Width           =   3975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Lash-whip and Bonesword 15pts/each"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   4095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "A pair of Boneswords 10pts/each"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Rending Claws 5pts/each"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox txtWarrior 
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
      Left            =   5040
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picWarrior 
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
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   7395
      TabIndex        =   2
      Top             =   720
      Width           =   7455
   End
   Begin VB.CommandButton cmdWarriorBack 
      Caption         =   "Back To Troops "
      Height          =   1335
      Left            =   1200
      TabIndex        =   0
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "The entire Brood may take:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   19
      Top             =   4560
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "The entire Brood may exchange its devourers for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   12
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "One Warrior may exchange its devourer for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Entire Brood can exchange its scything talons for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "3 - 9 per Brood:  30pts/Each "
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
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Tyranid Warrior Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmWarrior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberWarrior = txtWarrior.Text
EquipWarrior = EquipWarrior + (RendingClaws * NumberWarrior)
End Sub

Private Sub Check10_Click()
NumberWarrior = txtWarrior.Text
EquipWarrior = EquipWarrior + (TWAdrenalGlands * NumberWarrior)
End Sub

Private Sub Check11_Click()
NumberWarrior = txtWarrior.Text
EquipWarrior = EquipWarrior + (TWToxinSacs * NumberWarrior)

End Sub

Private Sub Check2_Click()
NumberWarrior = txtWarrior.Text
EquipWarrior = EquipWarrior + (BoneSwoard * NumberWarrior)
End Sub

Private Sub Check3_Click()
NumberWarrior = txtWarrior.Text
EquipWarrior = EquipWarrior + (AcidBlood * NumberWarrior)
End Sub

Private Sub Check4_Click()
NumberWarrior = txtWarrior.Text
EquipWarrior = EquipWarrior + BarbedStrangler
End Sub

Private Sub Check5_Click()
NumberWarrior = txtWarrior.Text
EquipWarrior = EquipWarrior + VCannon
End Sub

Private Sub Check8_Click()
NumberWarrior = txtWarrior.Text
EquipWarrior = EquipWarrior + (Deathspitter * NumberWarrior)

End Sub

Private Sub cmdCheckBoxCls_Click()
Check1 = 0
Check2 = 0
Check3 = 0
Check4 = 0
Check5 = 0
Check6 = 0
Check7 = 0
Check8 = 0
Check9 = 0
Check10 = 0
Check11 = 0

End Sub

Private Sub cmdNumberWarriorCls_Click()
txtWarrior = 0

End Sub

Private Sub cmdWarriorBack_Click()
frmWarrior.Hide
frmTroops.Show
picWarrior.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Warrior" Then
            picWarrior.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picWarrior.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picWarrior.Print "Sorry."
    End If

    
End Sub

Private Sub cmdWarriorTotal_Click()
NumberWarrior = txtWarrior.Text
WarriorTotal = WarriorTotal + (30 * NumberWarrior)

MsgBox "Your Warrior's are worth " & (WarriorTotal + EquipWarrior) & " pts "

End Sub

Private Sub cmdWarriorTotalCls_Click()
WarriorTotal = 0
EquipWarrior = 0

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\Warrior.JPG")

End Sub

