VERSION 5.00
Begin VB.Form frmHarpy 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Harpy"
   ClientHeight    =   12420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19920
   LinkTopic       =   "Form2"
   ScaleHeight     =   12420
   ScaleWidth      =   19920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHarpyTotal 
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
      TabIndex        =   17
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdHarpyTotalCls 
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
      TabIndex        =   16
      Top             =   5880
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
      TabIndex        =   15
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdHarpyCls 
      Caption         =   "Clear Number of Harpys"
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
      TabIndex        =   14
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Regeneration 15pts"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5280
      Width           =   3495
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Toxin Sacs 10tps"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Adrenal Glands 10pts"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Cluster Spines Free"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   3495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Twin-linked Heavy Venom Cannon 10pts"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox txtHarpy 
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
      Left            =   1800
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picHarpy 
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
      ScaleHeight     =   675
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   720
      Width           =   5175
   End
   Begin VB.CommandButton cmdHarpyBack 
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
      Left            =   5640
      TabIndex        =   0
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   $"frmHarpy.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3720
      TabIndex        =   13
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "A Harpy may take:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Replace Stinger Salvo with:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Replace Twin-Linked Stranglethorn Cannon for:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "1 per Brood:    160 pts / Each"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Harpy"
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
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmHarpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberHarpy = txtHarpy.Text
EquipHarpy = EquipHarpy + (TLHVCannon * NumberHarpy)

End Sub


Private Sub Check3_Click()
NumberHarpy = txtHarpy.Text
EquipHarpy = EquipHarpy + (AdrenalGlands * NumberHarpy)

End Sub

Private Sub Check4_Click()
NumberHarpy = txtHarpy.Text
EquipHarpy = EquipHarpy + (ToxinSacs * NumberHarpy)

End Sub

Private Sub Check5_Click()
NumberHarpy = txtHarpy.Text
EquipHarpy = EquipHarpy + (HarpyRegeneration * NumberHarpy)

End Sub

Private Sub cmdCheckBoxCls_Click()
Check1 = 0
Check2 = 0
Check3 = 0
Check4 = 0
Check5 = 0

End Sub

Private Sub cmdHarpyBack_Click()
frmHarpy.Hide
frmFastAttack.Show
picHarpy.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Harpy" Then
            picHarpy.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picHarpy.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picHarpy.Print "Sorry."
    End If

    
End Sub

Private Sub cmdHarpyCls_Click()
txtHarpy = 0

End Sub

Private Sub cmdHarpyTotal_Click()
NumberHarpy = txtHarpy.Text
HarpyTotal = HarpyTotal + (160 * NumberHarpy)

MsgBox "Your Total Points Spent on Harpys is " & (HarpyTotal + EquipHarpy)

End Sub

Private Sub cmdHarpyTotalCls_Click()
HarpyTotal = 0
EquipHarpy = 0
txtHarpy = 0


End Sub
