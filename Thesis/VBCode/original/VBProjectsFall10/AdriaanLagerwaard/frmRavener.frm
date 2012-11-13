VERSION 5.00
Begin VB.Form frmRavener 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ravener Brood"
   ClientHeight    =   12270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19425
   LinkTopic       =   "Form2"
   ScaleHeight     =   12270
   ScaleWidth      =   19425
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   4095
      Left            =   7080
      ScaleHeight     =   4035
      ScaleWidth      =   5355
      TabIndex        =   16
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton cmdRavenerTotal 
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
      TabIndex        =   15
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdRavenerTotalCls 
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
      TabIndex        =   14
      Top             =   5160
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
      TabIndex        =   13
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdNumberRaneverCls 
      Caption         =   "Clear Number of Raveners"
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
      TabIndex        =   12
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Deathspitters 10pts/each"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Devourers 5pts/each"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Spinefists 5pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Rending Claws 5pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3375
   End
   Begin VB.PictureBox picRavener 
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
      ScaleHeight     =   915
      ScaleWidth      =   5475
      TabIndex        =   3
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox txtRavener 
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
      Left            =   2880
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRavenerBack 
      Caption         =   "Back To Fast Attack"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      TabIndex        =   0
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Equiped With:         -Reinforced Chitin   -Scything Talons (two sets)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3840
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "The entire Brood may take one of the following:"
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
      TabIndex        =   7
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "The entire Brood may replace a set of Scything Talons for :"
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
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "3-9 per Brood:        30 pts/Each"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Ravener Brood"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmRavener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberRavener = txtRavener.Text
EquipRavener = EquipRavener + (RendingClaws * NumberRavener)

End Sub

Private Sub Check2_Click()
NumberRavener = txtRavener.Text
EquipRavener = EquipRavener + (SpineFists * NumberRavener)
End Sub

Private Sub Check3_Click()
NumberRavener = txtRavener.Text
EquipRavener = EquipRavener + (Devourer * NumberRavener)

End Sub

Private Sub Check4_Click()
NumberRavener = txtRavener.Text
EquipRavener = EquipRavener + (RevDeathspitter * NumberRavener)
End Sub

Private Sub cmdCheckBoxCls_Click()
Check1 = 0
Check2 = 0
Check3 = 0
Check4 = 0

End Sub

Private Sub cmdNumberRaneverCls_Click()
txtRavener = 0

End Sub

Private Sub cmdRavenerBack_Click()
frmRavener.Hide
frmFastAttack.Show
picRavener.Cls

End Sub

Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Ravener" Then
            picRavener.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picRavener.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picRavener.Print "Sorry."
    End If

    
End Sub

Private Sub cmdRavenerTotal_Click()
NumberRavener = txtRavener.Text
RavenerTotal = RavenerTotal + (30 * NumberRavener)

MsgBox "Your Total Points Spent on Raveners is " & (RavenerTotal + EquipRavener)

End Sub

Private Sub cmdRavenerTotalCls_Click()
RavenerTotal = 0
EquipRavener = 0
txtRavener = 0

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\ravener.GIF")

End Sub
