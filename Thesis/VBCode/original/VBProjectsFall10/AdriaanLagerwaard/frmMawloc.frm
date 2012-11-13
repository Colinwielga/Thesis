VERSION 5.00
Begin VB.Form frmMawloc 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Mawloc"
   ClientHeight    =   10890
   ClientLeft      =   4080
   ClientTop       =   2415
   ClientWidth     =   18375
   LinkTopic       =   "Form2"
   ScaleHeight     =   10890
   ScaleWidth      =   18375
   Begin VB.PictureBox picPicture 
      Height          =   5055
      Left            =   7560
      ScaleHeight     =   4995
      ScaleWidth      =   5235
      TabIndex        =   13
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton cmdMawlocTotal 
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
      TabIndex        =   12
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdMawlocTotalCls 
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
      TabIndex        =   11
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdMawlocCls 
      Caption         =   "Clear Number of Mawlocs and Check Boxes"
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
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Regeneration 25 pts"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Toxin Sacs 10 pts "
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Adrenal Glands 10pts"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
   End
   Begin VB.PictureBox picMawloc 
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
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txtMawloc 
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
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdMawlocBack 
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
      Left            =   5040
      TabIndex        =   0
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Equiped With:     -Bonded Exoskeleton -Claws and Teeth"
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
      TabIndex        =   9
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
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "1 per Brood: 170 pts/Each"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Mawloc"
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
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmMawloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberMawloc = txtMawloc
EquipMawloc = EquipMawloc + (AdrenalGlands * NumberMawloc)
End Sub

Private Sub Check2_Click()
NumberMawloc = txtMawloc
EquipMawloc = EquipMawloc + (ToxinSacs * NumberMawloc)

End Sub

Private Sub Check3_Click()
NumberMawloc = txtMawloc
EquipMawloc = EquipMawloc + (Regeneration * NumberMawloc)

End Sub

Private Sub cmdMawlocBack_Click()
frmMawloc.Hide
frmHeavySupport.Show
picMawloc.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Mawloc" Then
            picMawloc.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picMawloc.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picMawloc.Print "Sorry."
    End If

    
End Sub


Private Sub cmdMawlocCls_Click()
txtMawloc = 0
 Check1 = 0
 Check2 = 0
 Check3 = 0
 
End Sub

Private Sub cmdMawlocTotal_Click()
NumberMawloc = txtMawloc
MawlocTotal = MawlocTotal + (170 * NumberMawloc)

MsgBox "Your Total Points Spent on Mawlocs is " & (MawlocTotal + EquipMawloc)

End Sub

Private Sub cmdMawlocTotalCls_Click()
MawlocTotal = 0
EquipMawloc = 0
txtMawloc = 0

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\mawloc.JPG")

End Sub
