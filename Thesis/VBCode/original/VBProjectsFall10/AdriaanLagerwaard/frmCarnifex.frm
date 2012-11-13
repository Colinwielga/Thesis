VERSION 5.00
Begin VB.Form frmCarnifex 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Carnifex"
   ClientHeight    =   12135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19710
   LinkTopic       =   "Form6"
   ScaleHeight     =   12135
   ScaleWidth      =   19710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   5055
      Left            =   7440
      ScaleHeight     =   4995
      ScaleWidth      =   5475
      TabIndex        =   22
      Top             =   1200
      Width           =   5535
   End
   Begin VB.CommandButton cmdCarnifexTotal 
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
      Left            =   4800
      TabIndex        =   21
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdCarnifexTotalCls 
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
      Left            =   2520
      TabIndex        =   20
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdCarnifex 
      Caption         =   "Clear Number of Carnifexs and Check Boxes"
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
      TabIndex        =   19
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Regeneration 25pts/each"
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Bio Plasma 20pts/each"
      Height          =   495
      Left            =   3120
      TabIndex        =   17
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Toxin Sacs 10pts/each"
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Adrenal Glands 10pts/each"
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Frag Spines 5pts/each"
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Heavy Venom Cannon 25pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   3255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Stranglethorn Cannon 20pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Twin-linked Devourers w/ Brainleech worms 15pts/ each"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Twin-linked Deathspitter 15tps/each"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Crushing Claws 25pts/each"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2655
   End
   Begin VB.PictureBox picCarnifex 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   840
      Width           =   5295
   End
   Begin VB.TextBox txtCarnifex 
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
      Left            =   1920
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCarnifexBack 
      Caption         =   "Back To Heavy Support"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   0
      Top             =   7800
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Take any of the following:"
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
      Left            =   3120
      TabIndex        =   13
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Take one of the following, replace any set of Scything Talons:"
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
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Replace any set of Scything Talons with:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Replace one set of Scythign Talons with:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "1-3 per Brood:        160 pts/Each"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Carnifex"
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
      Width           =   1575
   End
End
Attribute VB_Name = "frmCarnifex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (CrushingClaws * NumberCarnifex)

End Sub

Private Sub Check10_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (Regeneration * NumberCarnifex)

End Sub

Private Sub Check2_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (TLDeathspitter * NumberCarnifex)

End Sub

Private Sub Check3_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (DevourerBLW * NumberCarnifex)

End Sub

Private Sub Check4_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (StrenglethornCannon * NumberCarnifex)

End Sub

Private Sub Check5_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (HVCannon * NumberCarnifex)

End Sub

Private Sub Check6_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (SpineFists * NumberCarnifex)

End Sub

Private Sub Check7_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (AdrenalGlands * NumberCarnifex)

End Sub

Private Sub Check8_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (ToxinSacs * NumberCarnifex)
End Sub

Private Sub Check9_Click()
NumberCarnifex = txtCarnifex.Text
EquipCarnifex = EquipCarnifex + (BioPlasma * NumberCarnifex)

End Sub

Private Sub cmdCarnifex_Click()
txtCarnifex = 0
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



End Sub

Private Sub cmdCarnifexBack_Click()
frmCarnifex.Hide
frmHeavySupport.Show
picCarnifex.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Carnifex" Then
            picCarnifex.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picCarnifex.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picCarnifex.Print "Sorry."
    End If

    
End Sub

Private Sub cmdCarnifexTotal_Click()
NumberCarnifex = txtCarnifex.Text
CarnifexTotal = CarnifexTotal + (160 * NumberCarnifex)

MsgBox "Total amount of points spent on Carnifex is " & (CarnifexTotal + EquipCarnifex)

End Sub

Private Sub cmdCarnifexTotalCls_Click()
CarnifexTotal = 0
EquipCarnifex = 0
txtCarnifex = 0

End Sub

Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\Carn.GIF")

End Sub
