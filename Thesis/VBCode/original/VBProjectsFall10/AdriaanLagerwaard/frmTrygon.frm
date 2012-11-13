VERSION 5.00
Begin VB.Form frmTrygon 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Trygon"
   ClientHeight    =   12075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19830
   LinkTopic       =   "Form3"
   ScaleHeight     =   12075
   ScaleWidth      =   19830
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   5535
      Left            =   7680
      ScaleHeight     =   5475
      ScaleWidth      =   5715
      TabIndex        =   14
      Top             =   480
      Width           =   5775
   End
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
      TabIndex        =   13
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
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
      TabIndex        =   12
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Number of Trygons and CheckBoxes"
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
      TabIndex        =   11
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Upgrade to Trygon Prime 40pts"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Regeneration 25pts"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Toxin Sacs 10pts"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Adrenal Glands 10pts"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   3015
   End
   Begin VB.PictureBox picTrygon 
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
      ScaleHeight     =   915
      ScaleWidth      =   4995
      TabIndex        =   4
      Top             =   720
      Width           =   5055
   End
   Begin VB.TextBox txtTrygon 
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
      Left            =   1800
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdTyrgonBack 
      Caption         =   "Back To Heavy Support"
      Height          =   1215
      Left            =   2280
      TabIndex        =   0
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   $"frmTrygon.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3720
      TabIndex        =   10
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label3 
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
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "1 per Brood:    200 pts/Each"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Trygon"
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
      Width           =   1335
   End
End
Attribute VB_Name = "frmTrygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberTrygon = txtTrygon
EquipTrygon = EquipTrygon + (AdrenalGlands * NumberTrygon)

End Sub

Private Sub Check2_Click()
NumberTrygon = txtTrygon
EquipTrygon = EquipTrygon + (ToxinSacs * NumberTrygon)

End Sub

Private Sub Check3_Click()
NumberTrygon = txtTrygon
EquipTrygon = EquipTrygon + (Regeneration * NumberTrygon)

End Sub

Private Sub Check4_Click()
NumberTrygon = txtTrygon
EquipTrygon = EquipTrygon + (TryPrime * NumberTrygon)

End Sub

Private Sub cmdTyrgonBack_Click()
frmTrygon.Hide
frmHeavySupport.Show
picTrygon.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Trygon" Then
            picTrygon.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picTrygon.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picTrygon.Print "Sorry."
    End If

    
End Sub


Private Sub Command1_Click()
txtTrygon = 0
Check1 = 0
Check2 = 0
Check3 = 0
Check4 = 0

End Sub

Private Sub Command2_Click()
TrygonTotal = 0
EquipTrygon = 0
txtTrygon = 0

End Sub

Private Sub Command3_Click()
NumberTrygon = txtTrygon
TrygonTotal = TrygonTotal + (200 * NumberTrygon)

MsgBox "Total Points Spent on Trygons is " & (TrygonTotal + EquipTrygon)

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\trygon.JPG")

End Sub
