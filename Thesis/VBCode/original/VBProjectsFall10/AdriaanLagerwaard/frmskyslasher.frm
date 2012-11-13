VERSION 5.00
Begin VB.Form frmSkySlasher 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sky Slasher Swarm"
   ClientHeight    =   12195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19695
   LinkTopic       =   "Form3"
   ScaleHeight     =   12195
   ScaleWidth      =   19695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSkySlasherTotal 
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
      Left            =   7080
      TabIndex        =   12
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdSkySlasherTotalCls 
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
      TabIndex        =   11
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheckBoxesCls 
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
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdNumberSkySlasherCls 
      Caption         =   "Clear Number of Sky Slasher Swarms"
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
      TabIndex        =   9
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Toxin Sacs 4pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Adrenal Glands 4pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Spinfists 5pts/each"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtSkySlasher 
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
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picSkySlasher 
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
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin VB.CommandButton cmdSkySlasherBack 
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
      Height          =   1095
      Left            =   5760
      TabIndex        =   0
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Equiped With:         -Chitin                       -Claws and Teeth    -Wings"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   1680
      Width           =   2535
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
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "3-9 per Brood:    15 pts/Each"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Sky Slasher Swarm"
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
Attribute VB_Name = "frmskyslasher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberSkySlasher = txtSkySlasher.Text
EquipSkySlasher = EquipSkySlasher + (SpineFists * NumberSkySlasher)

End Sub

Private Sub Check2_Click()
NumberSkySlasher = txtSkySlasher.Text
EquipSkySlasher = EquipSkySlasher + (RippersAdrenalGlands * NumberSkySlasher)

End Sub

Private Sub Check3_Click()
NumberSkySlasher = txtSkySlasher.Text
EquipSkySlasher = EquipSkySlasher + (RippersToxinSacs * NumberSkySlasher)

End Sub

Private Sub cmdCheckBoxesCls_Click()
Check1 = 0
Check2 = 0
Check3 = 0

End Sub

Private Sub cmdNumberSkySlasherCls_Click()
txtSkySlasher = 0

End Sub

Private Sub cmdSkySlasherBack_Click()
frmskyslasher.Hide
frmFastAttack.Show
picSkySlasher.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "SkySlasher" Then
            picSkySlasher.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picSkySlasher.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picSkySlasher.Print "Sorry."
    End If

    
End Sub

Private Sub cmdSkySlasherTotal_Click()
NumberSkySlasher = txtSkySlasher.Text
SkySlasherTotal = SkySlasherTotal + (15 * NumberSkySlasher)

MsgBox "Your Total Points Spent on Sky-Slashers is " & (SkySlasherTotal + EquipSkySlasher)

End Sub

Private Sub cmdSkySlasherTotalCls_Click()
SkySlasherTotal = 0
EquipSkySlasher = 0
txtSkySlasher = 0

End Sub
