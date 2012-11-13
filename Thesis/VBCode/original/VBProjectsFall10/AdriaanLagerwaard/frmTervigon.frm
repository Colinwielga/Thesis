VERSION 5.00
Begin VB.Form frmTervigon 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tervigon"
   ClientHeight    =   12405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   12405
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTervPoints 
      Caption         =   "Total Tervigon Points"
      Height          =   1215
      Left            =   6360
      TabIndex        =   22
      Top             =   7440
      Width           =   3135
   End
   Begin VB.CommandButton cmdTervCheckBoxCls 
      Caption         =   "Clear Check Boxes"
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton cmdTervPointCls 
      Caption         =   "Clear Tervigon Point Total"
      Height          =   1215
      Left            =   3120
      TabIndex        =   20
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Onslaught 15pts"
      Height          =   615
      Left            =   4080
      TabIndex        =   19
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Catalyst 15pts"
      Height          =   735
      Left            =   4080
      TabIndex        =   18
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Cluster Spines Free"
      Height          =   615
      Left            =   4080
      TabIndex        =   16
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Regeneration 30pts"
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   6600
      Width           =   3375
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Toxic Maisma 15pts"
      Height          =   735
      Left            =   360
      TabIndex        =   13
      Top             =   6120
      Width           =   3375
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Implant Attack 15pts"
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   5640
      Width           =   3375
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Acid Blood 15pts"
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   5160
      Width           =   3375
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Toxin Sacs 10pts"
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   4680
      Width           =   3375
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Adrenal Glands 10pts"
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   4200
      Width           =   3375
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Crushing Claws 25pts"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Scything Talons 2pts"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   5175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "160pts"
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
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.PictureBox picTervigon 
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
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   1080
      Width           =   6855
   End
   Begin VB.CommandButton cmdTervigonBack 
      Caption         =   "Go back to HQ Choices"
      Height          =   1215
      Left            =   7080
      TabIndex        =   0
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "Take any of the additional psychic powers:"
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
      TabIndex        =   17
      Top             =   3840
      Width           =   5415
   End
   Begin VB.Label Label5 
      Caption         =   "Replace stinger salvo with: "
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
      Left            =   4080
      TabIndex        =   15
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label4 
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
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Take one of the following:"
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
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Equiped With:             -Bonded Exoskeleton -Crushing Claws         Psychic Powers:          -Dominion"
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
      Left            =   7920
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Tervigon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmTervigon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value


Private Sub Check1_Click()
TervigonTotal = TervigonTotal + 160

End Sub

Private Sub Check11_Click()
TervigonTotal = TervigonTotal + Catalyst

End Sub

Private Sub Check12_Click()
TervigonTotal = TervigonTotal + Onslaught

End Sub

Private Sub Check2_Click()
TervigonTotal = TervigonTotal + GeneScythingTalons

End Sub

Private Sub Check3_Click()
TervigonTotal = TervigonTotal + CrushingClaws
End Sub

Private Sub Check4_Click()
TervigonTotal = TervigonTotal + AdrenalGlands
End Sub

Private Sub Check5_Click()
TervigonTotal = TervigonTotal + ToxinSacs

End Sub

Private Sub Check6_Click()
TervigonTotal = TervigonTotal + AcidBlood

End Sub

Private Sub Check7_Click()
TervigonTotal = TervigonTotal + ImplantAttack

End Sub

Private Sub Check8_Click()
TervigonTotal = TervigonTotal + ToxicMaisma

End Sub

Private Sub Check9_Click()
TervigonTotal = TervigonTotal + TervRegeneration

End Sub

Private Sub cmdTervCheckBoxCls_Click()
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
Check12 = 0
TervigonTotal = TervigonTotal / 2


End Sub

Private Sub cmdTervigonBack_Click()
frmTervigon.Hide
frmHQ.Show
picTervigon.Cls



End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean

 For j = 1 To CTR
        If names(j) = "Tervigon" Then
            picTervigon.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picTervigon.Print " "; WS(j); "       "; BS(j); "      "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
Next j
    
    If Not found Then
        picTervigon.Print "Sorry."
    End If
    
End Sub

Private Sub cmdTervPointCls_Click()
TervigonTotal = 0

End Sub

Private Sub cmdTervPoints_Click()
MsgBox "Total Points spent on Tervigon. " & TervigonTotal

End Sub
