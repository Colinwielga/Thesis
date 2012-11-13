VERSION 5.00
Begin VB.Form frmTyrannofex 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tyrannofex"
   ClientHeight    =   12090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19620
   LinkTopic       =   "Form1"
   ScaleHeight     =   12090
   ScaleWidth      =   19620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      TabIndex        =   19
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdTyranTotalCls 
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
      TabIndex        =   18
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdTyranCls 
      Caption         =   "Clear Number of Tyrannofexs and Check Boxes"
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
      TabIndex        =   17
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Regeneration 30pts"
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Toxin Sacs 10pts"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Adrenal Glands 10pts"
      Height          =   615
      Left            =   3360
      TabIndex        =   14
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Shreddershard Beetles Free"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   3015
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Diseccator Larvae Free"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   3015
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Cluster Spines Free"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Rupture Cannon 15pts"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Fleshborer hive 10pts"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2895
   End
   Begin VB.PictureBox picTyran 
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
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox txtTyran 
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
      Left            =   2160
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdTyrannofexBack 
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
      Height          =   1215
      Left            =   5160
      TabIndex        =   0
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   $"frmTyrannofex.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   6360
      TabIndex        =   20
      Top             =   1800
      Width           =   2895
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
      Left            =   3360
      TabIndex        =   13
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Replace Electroshock Grubs with:"
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
      TabIndex        =   10
      Top             =   3960
      Width           =   3015
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
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Replace Acid Spray with:"
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
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "1 per Brood:      250 pts/Each"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Tyrannofex"
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
      Width           =   1815
   End
End
Attribute VB_Name = "frmTyrannofex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberTyran = txtTyran.Text
EquipTyran = EquipTyran + (BarbedStrangler * NumberTyran)

End Sub

Private Sub Check2_Click()
NumberTyran = txtTyran.Text
EquipTyran = EquipTyran + (RuptureCannon * NumberTyran)

End Sub

Private Sub Check6_Click()
NumberTyran = txtTyran.Text
EquipTyran = EquipTyran + (AdrenalGlands * NumberTyran)

End Sub

Private Sub Check7_Click()
NumberTyran = txtTyran.Text
EquipTyran = EquipTyran + (ToxinSacs * NumberTyran)

End Sub

Private Sub Check8_Click()
NumberTyran = txtTyran.Text
EquipTyran = EquipTyran + (TyranRegeneration * NumberTyran)

End Sub

Private Sub cmdTyranCls_Click()
txtTyran = 0
Check1 = 0
Check2 = 0
Check3 = 0
Check4 = 0
Check5 = 0
Check6 = 0
Check7 = 0
Check8 = 0

End Sub

Private Sub cmdTyrannofexBack_Click()
frmTyrannofex.Hide
frmHeavySupport.Show
picTyran.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Tyrannofex" Then
            picTyran.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picTyran.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picTyran.Print "Sorry."
    End If

    
End Sub


Private Sub cmdTyranTotalCls_Click()
TyranTotal = 0
EquipTyran = 0
txtTyran = 0

End Sub

Private Sub Command3_Click()
NumberTyran = txtTyran.Text
TyranTotal = TyranTotal + (250 * NumberTyran)

MsgBox " Total Points spent on Tyrannofex is " & (TyranTotal + EquipTyran)

End Sub
