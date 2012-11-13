VERSION 5.00
Begin VB.Form frmGargoyle 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gargoyle Brood"
   ClientHeight    =   12540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19980
   LinkTopic       =   "Form1"
   ScaleHeight     =   12540
   ScaleWidth      =   19980
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   3615
      Left            =   10440
      ScaleHeight     =   3555
      ScaleWidth      =   4635
      TabIndex        =   13
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton cmdGargTotal 
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
      Left            =   6960
      TabIndex        =   12
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdGargTotalCls 
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
      Left            =   4680
      TabIndex        =   11
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheckBoxCls 
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
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdGargCls 
      Caption         =   "Clear Number of Gargoyles"
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
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Toxic Sacs 1pt/each"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Adrenal Glands 1pt/each"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtGarg 
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
      Left            =   3120
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picGarg 
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
      Left            =   240
      ScaleHeight     =   795
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin VB.CommandButton cmdGargoyleBack 
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
      Height          =   1335
      Left            =   5520
      TabIndex        =   0
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Equiped With:     -Blinding Venom -Chitin                  -Claws and Teeth -Fleshborer            -Wings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3480
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
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
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "10-30 per Brood:  6 pts/Each"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Gargoyle Brood"
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
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmGargoyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
NumberGarg = txtGarg.Text
EquipGarg = EquipGarg + (TermAdrenalGlands * NumberGarg)

End Sub

Private Sub Check2_Click()
NumberGarg = txtGarg.Text
EquipGarg = EquipGarg + (TermToxinSacs * NumberGarg)

End Sub

Private Sub cmdCheckBoxCls_Click()
Check1 = 0
Check2 = 0

End Sub

Private Sub cmdGargCls_Click()
txtGarg = 0

End Sub

Private Sub cmdGargoyleBack_Click()
frmGargoyle.Hide
frmFastAttack.Show
picGarg.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Gargoyle" Then
            picGarg.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picGarg.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picGarg.Print "Sorry."
    End If

    
End Sub

Private Sub cmdGargTotal_Click()
NumberGarg = txtGarg.Text
GargTotal = GargTotal + (6 * NumberGarg)

MsgBox "Your Total Points Spent on Gargoyles is " & (GargTotal + EquipGarg)

End Sub

Private Sub cmdGargTotalCls_Click()
GargTotal = 0
EquipGarg = 0
txtGarg = 0

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\garg.JPG")

End Sub
