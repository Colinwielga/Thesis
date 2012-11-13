VERSION 5.00
Begin VB.Form frmSwarmLord 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The SwarmLord"
   ClientHeight    =   12045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15795
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12045
   ScaleWidth      =   15795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwarmLordClearCheckbox 
      Caption         =   "Clear check box"
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
      Left            =   4200
      TabIndex        =   7
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton cmdSwarmLord 
      Caption         =   "Clear SwarmLord points "
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
      Left            =   360
      TabIndex        =   6
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "280pts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.PictureBox picSwarmLord 
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
      ScaleWidth      =   7515
      TabIndex        =   2
      Top             =   1560
      Width           =   7575
   End
   Begin VB.CommandButton cmdSwarmLordBack 
      Caption         =   "Go back to HQ Choices"
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
      Left            =   360
      TabIndex        =   0
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Psychic Powers: -The Horror         -Psychic Scream -Paroxysm            -Leech Essence"
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
      Left            =   4800
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Equiped with:                  -Bonded Exoskeleton     -Bonesabers"
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
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label lblSwarmLord 
      Caption         =   "The SwarmLord"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmSwarmLord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms


Private Sub Check1_Click()
SwarmLordTotal = SwarmLordTotal + 280
End Sub

Private Sub cmdSwarmLord_Click()
SwarmLordTotal = 0




End Sub

Private Sub cmdSwarmLordBack_Click()

frmSwarmLord.Hide
frmHQ.Show
picSwarmLord.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean
 For j = 1 To CTR
        If names(j) = "TheSwarmLord" Then
            picSwarmLord.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picSwarmLord.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picSwarmLord.Print "Sorry."
    End If

    
End Sub

Private Sub cmdSwarmLordClearCheckbox_Click()
Check1 = 0
SwarmLordTotal = SwarmLordTotal - 280
End Sub
