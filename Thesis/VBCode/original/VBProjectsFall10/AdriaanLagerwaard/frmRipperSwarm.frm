VERSION 5.00
Begin VB.Form frmRipperSwarm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ripper Swarm Brood"
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form2"
   ScaleHeight     =   11940
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNumberRipperSwarm 
      Caption         =   "Clear Number of Ripper Swarms"
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
      Left            =   7920
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdRipperSwarmTotal 
      Caption         =   "Total Points "
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
      Left            =   3840
      TabIndex        =   13
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdRipperSwarmTotalCls 
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
      Left            =   3840
      TabIndex        =   12
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdRipperSwarmCheckBoxCls 
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
      Left            =   3840
      TabIndex        =   11
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Tunnel Swarm 2tps/Each"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   3375
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Toxin Sacs 4pts/Each"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Adrenal Glands 4pts/Each"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Spinfists 5pts/Each"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3375
   End
   Begin VB.PictureBox picRipperSwarm 
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
      TabIndex        =   4
      Top             =   1080
      Width           =   5535
   End
   Begin VB.TextBox txtRipperSwarm 
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
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdRipperSwarmBack 
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
      Left            =   360
      TabIndex        =   0
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Equiped With:           -Chitin                        -Claws and Teeth"
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
      Left            =   3960
      TabIndex        =   10
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "The Entire Brood may take:"
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
      TabIndex        =   9
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "3-9 per Brood, up to 6 Broods: 10pts/ Each"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Ripper Swarm Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmRipperSwarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Public Sub loadData()
Dim j As Integer, found As Boolean






 For j = 1 To CTR
        If names(j) = "RipperSwarm" Then
            picRipperSwarm.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picRipperSwarm.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picRipperSwarm.Print "Sorry."
    End If

    
End Sub

Private Sub Check1_Click()
NumberRipperSwarm = txtRipperSwarm.Text
EquipRipperSwarm = EquipRipperSwarm + (SpineFists * NumberRipperSwarm)
End Sub

Private Sub Check2_Click()
NumberRipperSwarm = txtRipperSwarm.Text
EquipRipperSwarm = EquipRipperSwarm + (RippersAdrenalGlands * NumberRipperSwarm)

End Sub

Private Sub Check3_Click()
NumberRipperSwarm = txtRipperSwarm.Text
EquipRipperSwarm = EquipRipperSwarm + (RippersToxinSacs * NumberRipperSwarm)
End Sub

Public Sub Check4_Click()
NumberRipperSwarm = txtRipperSwarm.Text
EquipRipperSwarm = EquipRipperSwarm + (TunnelSwarm * NumberRipperSwarm)
End Sub

Private Sub cmdNumberRipperSwarm_Click()
txtRipperSwarm = 0

End Sub

Private Sub cmdRipperSwarmBack_Click()
frmRipperSwarm.Hide
frmTroops.Show
picRipperSwarm.Cls

End Sub


Private Sub cmdRipperSwarmCheckBoxCls_Click()

Check1 = 0
Check2 = 0
Check3 = 0
Check4 = 0

'EquipRipperSwarm = EquipRipperSwarm / 2

End Sub

Public Sub cmdRipperSwarmTotal_Click()

NumberRipperSwarm = txtRipperSwarm.Text
RipperSwarmTotal = RipperSwarmTotal + (10 * NumberRipperSwarm)

MsgBox "You have " & NumberRipperSwarm & " worth " & (RipperSwarmTotal + EquipRipperSwarm)





End Sub

Private Sub cmdRipperSwarmTotalCls_Click()
RipperSwarmTotal = 0
EquipRipperSwarm = 0


End Sub

