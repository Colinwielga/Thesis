VERSION 5.00
Begin VB.Form frmHormagaunt 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hormagaunt Brood"
   ClientHeight    =   11850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19035
   LinkTopic       =   "Form3"
   ScaleHeight     =   11850
   ScaleWidth      =   19035
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   3975
      Left            =   9480
      ScaleHeight     =   3915
      ScaleWidth      =   4875
      TabIndex        =   13
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton cmdHormTotal 
      Caption         =   "Total Points Spent"
      Height          =   975
      Left            =   6960
      TabIndex        =   12
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdHormToralCls 
      Caption         =   "Clear Points Spent"
      Height          =   975
      Left            =   4680
      TabIndex        =   11
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheckBoxCls 
      Caption         =   "Clear Check Boxes"
      Height          =   975
      Left            =   2400
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdNumberHormCls 
      Caption         =   "Clear Number of Hormagaunts"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Toxic Sacs 2pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Adrenal Glands 2pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtHorm 
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
      Left            =   3360
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picHorm 
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
      ScaleHeight     =   675
      ScaleWidth      =   5355
      TabIndex        =   2
      Top             =   720
      Width           =   5415
   End
   Begin VB.CommandButton cmdHormagauntBack 
      Caption         =   "Back To Troops"
      Height          =   1095
      Left            =   6600
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Equiped With:              -Chitin                           -Scything Talons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label3 
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
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "10-30 per Brood: 6 pts/Each"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Hormagaunt Brood"
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
      Width           =   3375
   End
End
Attribute VB_Name = "frmHormagaunt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberHorm = txtHorm.Text
EquipHorm = EquipHorm + (HormAdrenalGlands * NumberHorm)

End Sub

Private Sub Check2_Click()
NumberHorm = txtHorm.Text
EquipHorm = EquipHorm + (HormToxinSacs * NumberHorm)

End Sub

Private Sub cmdCheckBoxCls_Click()
Check1 = 0
Check2 = 0

End Sub

Private Sub cmdHormagauntBack_Click()
frmHormagaunt.Hide
frmTroops.Show
picHorm.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Hormagaunt" Then
            picHorm.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picHorm.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picHorm.Print "Sorry."
    End If

    
End Sub


Private Sub cmdHormToralCls_Click()
HormTotal = 0
EquipHorm = 0
txtHorm = 0


End Sub

Private Sub cmdHormTotal_Click()
NumberHorm = txtHorm.Text
HormTotal = HormTotal + (6 * NumberHorm)

MsgBox "Your Total Points Spent on Hormagaunts is " & (HormTotal + EquipHorm)

End Sub

Private Sub cmdNumberHormCls_Click()
txtHorm = 0

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\horm.JPG")

End Sub
