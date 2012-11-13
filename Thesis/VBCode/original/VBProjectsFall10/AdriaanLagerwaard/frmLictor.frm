VERSION 5.00
Begin VB.Form frmLictor 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Lictor Brood"
   ClientHeight    =   12480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19860
   LinkTopic       =   "Form7"
   ScaleHeight     =   12480
   ScaleWidth      =   19860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   4695
      Left            =   6840
      ScaleHeight     =   4635
      ScaleWidth      =   4275
      TabIndex        =   9
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdLictorTotal 
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
      Left            =   3240
      TabIndex        =   8
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton cmdLictorTotalCls 
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
      Left            =   3240
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdNumberLictorCls 
      Caption         =   "Clear Number of Lictors"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtLictor 
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
      Left            =   2760
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picLictor 
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
      TabIndex        =   2
      Top             =   840
      Width           =   5535
   End
   Begin VB.CommandButton cmdLictorBack 
      Caption         =   "Back to Elites"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "1-9, up to 3 per Brood: 65pts/Each"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Equiped With:         -Chameleonic Skin -Flesh Hooks           -Reinforced Chitin   -Rending Claws       -Scything Talons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Lictor Brood"
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
      Width           =   2175
   End
End
Attribute VB_Name = "frmLictor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdLictorBack_Click()
frmLictor.Hide
frmElites.Show
picLictor.Cls

End Sub


Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Lictor" Then
            picLictor.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A       "; "Ld     "; "Sv      "
            picLictor.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picLictor.Print "Sorry."
    End If

    
End Sub



Private Sub cmdLictorTotalCls_Click()
LictorTotal = 0

End Sub

Private Sub cmdNumberLictorCls_Click()
txtLictor = 0
LictorTotal = LictorTotal - (50 * NumberLictor)
End Sub

Private Sub cmdLictorTotal_Click()
Dim NumberLictor As Single
NumberLictor = txtLictor.Text
LictorTotal = LictorTotal + (65 * NumberLictor)

MsgBox "You have " & NumberLictor & " worth. " & LictorTotal

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\Lictor.JPG")

End Sub
