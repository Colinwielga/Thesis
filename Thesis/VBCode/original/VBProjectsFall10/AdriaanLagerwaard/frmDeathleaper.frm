VERSION 5.00
Begin VB.Form frmDeathleaper 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Deathleaper"
   ClientHeight    =   12780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20175
   LinkTopic       =   "Form6"
   ScaleHeight     =   12780
   ScaleWidth      =   20175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDeathleaperTotal 
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
      Left            =   3360
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton cmdDeathleaperTotalCls 
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
      Left            =   3360
      TabIndex        =   6
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton cmdDeathleaperCheckBoxCls 
      Caption         =   "Clear Check Box"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "140pts"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picDeathleaper 
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
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   600
      Width           =   5175
   End
   Begin VB.CommandButton cmdDeathleaperBack 
      Caption         =   "Back To Elites"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Equiped With:      -Chameleonic Skin   -Flesh Hooks            -Reinforced Chitin     -Rending Claws        -Scything Talons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Deathleaper"
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
Attribute VB_Name = "frmDeathleaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms


Private Sub Check1_Click()
DeathleaperTotal = DeathleaperTotal + 140

End Sub

Private Sub cmdDeathleaperBack_Click()
frmDeathleaper.Hide
frmElites.Show
picDeathleaper.Cls

End Sub
Public Sub loadData()



Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Deathleaper" Then
            picDeathleaper.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picDeathleaper.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picDeathleaper.Print "Sorry."
    End If
End Sub


Private Sub cmdDeathleaperCheckBoxCls_Click()
Check1 = 0
DeathleaperTotal = DeathleaperTotal / 2

End Sub

Private Sub cmdDeathleaperTotal_Click()
MsgBox "Points spent on Deathleaper " & DeathleaperTotal

End Sub

Private Sub cmdDeathleaperTotalCls_Click()
DeathleaperTotal = 0

End Sub
