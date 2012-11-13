VERSION 5.00
Begin VB.Form frmHiveTyrant 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hive Tyrant"
   ClientHeight    =   13815
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   13815
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   5535
      Left            =   12360
      ScaleHeight     =   5475
      ScaleWidth      =   6315
      TabIndex        =   30
      Top             =   2520
      Width           =   6375
   End
   Begin VB.CommandButton cmdChecksCls 
      BackColor       =   &H000000FF&
      Caption         =   "Clear Checks"
      Height          =   1095
      Left            =   840
      MaskColor       =   &H000040C0&
      TabIndex        =   29
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   2775
   End
   Begin VB.CommandButton cmdHTClear 
      Caption         =   "Clear HiveTyrant Points"
      Height          =   1095
      Left            =   4680
      TabIndex        =   28
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdHTPoints 
      BackColor       =   &H000080FF&
      Caption         =   "Hive Tyrant Points Total"
      Height          =   1095
      Left            =   8280
      TabIndex        =   27
      Top             =   7920
      Width           =   4095
   End
   Begin VB.CheckBox Check18 
      Caption         =   "170pts"
      Height          =   495
      Left            =   2760
      TabIndex        =   26
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox Check17 
      Caption         =   "wings 60pts"
      Height          =   615
      Left            =   7800
      TabIndex        =   25
      Top             =   6840
      Width           =   4575
   End
   Begin VB.CheckBox Check16 
      Caption         =   "Armoured shell 40pts"
      Height          =   495
      Left            =   7800
      TabIndex        =   24
      Top             =   6480
      Width           =   4575
   End
   Begin VB.CheckBox Check15 
      Caption         =   "Thorax swarm with either electrochock grubs, disiccator larvae or shreddershard beetles 25pts"
      Height          =   615
      Left            =   7800
      TabIndex        =   23
      Top             =   6000
      Width           =   4575
   End
   Begin VB.CheckBox Check14 
      Caption         =   "Regeneration 20pts"
      Height          =   495
      Left            =   7800
      TabIndex        =   21
      Top             =   5040
      Width           =   5295
   End
   Begin VB.CheckBox Check13 
      Caption         =   "Toxic miasma 15tps"
      Height          =   375
      Left            =   7800
      TabIndex        =   20
      Top             =   4680
      Width           =   5295
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Implant attack 15pts"
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   4320
      Width           =   5295
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Acid blood 15pts"
      Height          =   495
      Left            =   7800
      TabIndex        =   18
      Top             =   3840
      Width           =   5055
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Toxin sacs 10pts"
      Height          =   375
      Left            =   7800
      TabIndex        =   17
      Top             =   3480
      Width           =   4575
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Adrenal Glands 10pts"
      Height          =   495
      Left            =   7800
      TabIndex        =   16
      Top             =   3000
      Width           =   4575
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Old Adversary 25tps"
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   7320
      Width           =   6735
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Indescribable Horror 25pts"
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   6840
      Width           =   6735
   End
   Begin VB.CheckBox Check6 
      Caption         =   "hive commander 25pts"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   6480
      Width           =   6735
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Heavy venom cannon 25pts"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   5640
      Width           =   6735
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Stranglethorn cannon 20pts"
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   5160
      Width           =   6735
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "An additinoal set of scyting talons free"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3120
      Width           =   6735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Twin-linked Devourers w/ brainleech worms 15pts"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   4320
      Width           =   6735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Twin-linked Deathspitter 15pts"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3960
      Width           =   6735
   End
   Begin VB.PictureBox picHiveTyrant 
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
      ScaleHeight     =   1155
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   1320
      Width           =   7335
   End
   Begin VB.CommandButton cmdHiveTyrantBack 
      Caption         =   "Go Back to HQ choices"
      Height          =   1215
      Left            =   8400
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label6 
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
      Height          =   615
      Left            =   7800
      TabIndex        =   22
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Label Label5 
      Caption         =   "Take any of the following"
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
      Left            =   7800
      TabIndex        =   15
      Top             =   2640
      Width           =   6615
   End
   Begin VB.Label Label4 
      Caption         =   "Take any of the following abilities:"
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
      Left            =   600
      TabIndex        =   12
      Top             =   6120
      Width           =   6735
   End
   Begin VB.Label Label3 
      Caption         =   "Take one of the following, replace any set of  scything talons:"
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
      Left            =   600
      TabIndex        =   8
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Label Label2 
      Caption         =   "Replace lash whip and boneswword with:"
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
      Left            =   600
      TabIndex        =   6
      Top             =   2640
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Replace any set of scything talons with:"
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
      Left            =   600
      TabIndex        =   3
      Top             =   3480
      Width           =   6735
   End
   Begin VB.Label lblHiveTyrant 
      Caption         =   "Hive Tyrant "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmHiveTyrant"
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
        If names(j) = "HiveTyrant" Then
            picHiveTyrant.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picHiveTyrant.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picHiveTyrant.Print "Sorry."
    End If

    
End Sub




Private Sub Check1_Click()
HiveTyrantTotal = HiveTyrantTotal + TLDeathspitter

End Sub

Private Sub Check10_Click()
HiveTyrantTotal = HiveTyrantTotal + ToxinSacs
End Sub

Private Sub Check11_Click()
HiveTyrantTotal = HiveTyrantTotal + AcidBlood

End Sub

Private Sub Check12_Click()
HiveTyrantTotal = HiveTyrantTotal + ImplantAttack
End Sub

Private Sub Check13_Click()
HiveTyrantTotal = HiveTyrantTotal + ToxicMaisma

End Sub

Private Sub Check14_Click()
HiveTyrantTotal = HiveTyrantTotal + HTRegeneration
End Sub

Private Sub Check15_Click()
HiveTyrantTotal = HiveTyrantTotal + ThoraxSwarm
End Sub

Private Sub Check16_Click()
HiveTyrantTotal = HiveTyrantTotal + ArmouredShell
End Sub

Private Sub Check17_Click()
HiveTyrantTotal = HiveTyrantTotal + Wings
End Sub

Private Sub Check18_Click()
HiveTyrantTotal = HiveTyrantTotal + 170
End Sub

Private Sub Check2_Click()
HiveTyrantTotal = HiveTyrantTotal + DevourerBLW
End Sub

Private Sub Check4_Click()
HiveTyrantTotal = HiveTyrantTotal + StrenglethornCannon
End Sub

Private Sub Check5_Click()
HiveTyrantTotal = HiveTyrantTotal + HVCannon


End Sub

Private Sub Check6_Click()
HiveTyrantTotal = HiveTyrantTotal + 25
End Sub

Private Sub Check7_Click()
HiveTyrantTotal = HiveTyrantTotal + 25

End Sub

Private Sub Check8_Click()
HiveTyrantTotal = HiveTyrantTotal + 25
End Sub

Private Sub Check9_Click()
HiveTyrantTotal = HiveTyrantTotal + AdrenalGlands

End Sub

Private Sub cmdChecksCls_Click()
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
Check13 = 0
Check14 = 0
Check15 = 0
Check16 = 0
Check17 = 0
Check18 = 0
HiveTyrantTotal = HiveTyrantTotal / 2

End Sub

Private Sub cmdHiveTyrantBack_Click()

frmHiveTyrant.Hide
frmHQ.Show
picHiveTyrant.Cls

End Sub


Private Sub cmdHTClear_Click()
HiveTyrantTotal = 0
End Sub

Private Sub cmdHTPoints_Click()
MsgBox "Total points spent on Hive Tyrant are " & HiveTyrantTotal
End Sub

Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\HiveTyrant.JPG")

End Sub
