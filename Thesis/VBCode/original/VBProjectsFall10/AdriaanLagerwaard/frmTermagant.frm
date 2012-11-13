VERSION 5.00
Begin VB.Form frmTermagant 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Termagant Brood"
   ClientHeight    =   11520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19110
   LinkTopic       =   "Form4"
   ScaleHeight     =   11520
   ScaleWidth      =   19110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   3015
      Left            =   7080
      ScaleHeight     =   2955
      ScaleWidth      =   3675
      TabIndex        =   17
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton cmdTermTotal 
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
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdTermTotalCls 
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
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheckBoxCls 
      Caption         =   "Clear Chech Boxes"
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
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdNumberTermCls 
      Caption         =   "Clear Number fo Termagants"
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
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Toxin Sacs 1pt/each"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Adrenal Glands 1pt/each"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Devourer 5pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   3375
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Spike Rifle 1pt/each"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Spinfists 1pt/each"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtTerm 
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
      Left            =   3240
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picTerm 
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
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   720
      Width           =   5535
   End
   Begin VB.CommandButton cmdTermagantBack 
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
      Left            =   6720
      TabIndex        =   0
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Equiped With:         -Chitin                       -Claws and Teeth    -Feshborer"
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
      Left            =   3960
      TabIndex        =   16
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "The entire Brood may repalace their fleshborers for:"
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
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "10-30 per Brood:    5 pts/Each"
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
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Termagant Brood"
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
      Width           =   3015
   End
End
Attribute VB_Name = "frmTermagant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check2_Click()
NumberTerm = txtTerm.Text
EquipTerm = EquipTerm + (TermSpineFists * NumberTerm)

End Sub

Private Sub Check3_Click()
NumberTerm = txtTerm.Text
EquipTerm = EquipTerm + (SpikeRifle * NumberTerm)
End Sub

Private Sub Check4_Click()
NumberTerm = txtTerm.Text
EquipTerm = EquipTerm + (Devourer * NumberTerm)
End Sub

Private Sub Check5_Click()
NumberTerm = txtTerm.Text
EquipTerm = EquipTerm + (TermAdrenalGlands * NumberTerm)
End Sub

Private Sub Check6_Click()
NumberTerm = txtTerm.Text
EquipTerm = EquipTerm + (TermToxinSacs * NumberTerm)

End Sub

Private Sub cmdCheckBoxCls_Click()

Check2 = 0
Check3 = 0
Check4 = 0
Check5 = 0
Check6 = 0
End Sub

Private Sub cmdNumberTermCls_Click()
txtTerm = 0

End Sub

Private Sub cmdTermagantBack_Click()
frmTermagant.Hide
frmTroops.Show
picTerm.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Termagant" Then
            picTerm.Print "WS     "; "BS       "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picTerm.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picTerm.Print "Sorry."
    End If

    
End Sub

Private Sub cmdTermTotal_Click()
NumberTerm = txtTerm.Text
TermTotal = TermTotal + (5 * NumberTerm)

MsgBox "Your Points Total for Termagants is " & (TermTotal + EquipTerm)

End Sub

Private Sub cmdTermTotalCls_Click()
TermTotal = 0
EquipTerm = 0

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\Term.GIF")

End Sub
