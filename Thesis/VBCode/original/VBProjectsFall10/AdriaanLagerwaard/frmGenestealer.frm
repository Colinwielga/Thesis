VERSION 5.00
Begin VB.Form frmGenestealer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Genestealer Brood"
   ClientHeight    =   11880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form5"
   ScaleHeight     =   11880
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   3735
      Left            =   10560
      ScaleHeight     =   3675
      ScaleWidth      =   3795
      TabIndex        =   20
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton cmdBroodlordCls 
      Caption         =   "Clear Broodlord Check Boxes"
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
      TabIndex        =   18
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txtGenestealer 
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
      Left            =   3600
      TabIndex        =   16
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdGenestealerTotal 
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
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenestealerTotalCls 
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
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton mcdCheckBoxesCls 
      Caption         =   "Clear Genestealer Check Boxes"
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
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdNumberGenestealerCls 
      Caption         =   "Clear Number of Genestealers"
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
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Toxin Sacs 3pts/each"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Adrenal Glands 3pts/each"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   2880
      Width           =   3375
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Scything Talons 2pts/each"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Acid Blood 15pts"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   3855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Implant Attack 15pts"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scything Talons 2pts"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Upgrade one Genestealer to a Broodlord 46pts"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   3735
   End
   Begin VB.PictureBox picGenestealer 
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
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   720
      Width           =   6255
   End
   Begin VB.CommandButton cmdGenestealerBack 
      Caption         =   "Back To Troops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6120
      TabIndex        =   0
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   $"frmGenestealer.frx":0000
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
      Left            =   7320
      TabIndex        =   19
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "5-10 per Brood:    14pts/ Each"
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
      Left            =   4440
      TabIndex        =   17
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "All Genestealers in the Brood may take:"
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
      Left            =   3960
      TabIndex        =   8
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "A Broodlord may take:"
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
      TabIndex        =   3
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Genestealer Brood"
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
      Width           =   3135
   End
End
Attribute VB_Name = "frmGenestealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
EquipBroodlord = EquipBroodlord + 46

End Sub

Private Sub Check2_Click()
EquipBroodlord = EquipBroodlord + GeneScythingTalons

End Sub

Private Sub Check3_Click()
EquipBroodlord = EquipBroodlord + ImplantAttack

End Sub

Private Sub Check4_Click()
EquipBroodlord = EquipBroodlord + AcidBlood

End Sub

Private Sub Check5_Click()
NumberGenestealer = txtGenestealer.Text
EquipGenestealer = EquipGenestealer + (GeneScythingTalons * NumberGenestealer)

End Sub

Private Sub Check6_Click()
NumberGenestealer = txtGenestealer.Text
EquipGenestealer = EquipGenestealer + (GeneAdrenalGlands * NumberGenestealer)

End Sub

Private Sub Check7_Click()
NumberGenestealer = txtGenestealer.Text
EquipGenestealer = EquipGenestealer + (GeneToxinSacs * NumberGenestealer)

End Sub

Private Sub cmdBroodlordCls_Click()
Check1 = 0
   
Check2 = 0
  
Check3 = 0
    
Check4 = 0
 EquipBroodlord = EquipBroodlord / 2
End Sub

Private Sub cmdGenestealerBack_Click()
frmGenestealer.Hide
frmTroops.Show
picGenestealer.Cls

End Sub
 
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Genestealer" Then
            picGenestealer.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picGenestealer.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j

    If Not found Then
        picGenestealer.Print "Sorry."
    End If

    
End Sub

Private Sub cmdGenestealerTotal_Click()
NumberGenestealer = txtGenestealer.Text
GenestealerTotal = GenestealerTotal + (14 * NumberGenestealer)

MsgBox "Your total points spent on Genestealers is " & (GenestealerTotal + EquipGenestealer + EquipBroodlord)

End Sub

Private Sub cmdGenestealerTotalCls_Click()
GenestealerTotal = 0
EquipGenestealer = 0

End Sub

Private Sub cmdNumberGenestealerCls_Click()
txtGenestealer = 0

End Sub

Private Sub mcdCheckBoxesCls_Click()
Check1 = 0
    'If Check1 = 0 Then
      '  EquipGenestealer = EquipGenestealer - 46
   ' End If
Check2 = 0
    If Check2 = 0 Then
        EquipGenestealer = EquipGenestealer - 2
    End If
Check3 = 0
    If Check3 = 0 Then
        EquipGenestealer = EquipGenestealer - 15
    End If
Check4 = 0
    If Check4 = 0 Then
        EquipGenestealer = EquipGenestealer - 15
    End If
Check5 = 0
Check6 = 0
Check7 = 0

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\Gene.JPG")

End Sub
