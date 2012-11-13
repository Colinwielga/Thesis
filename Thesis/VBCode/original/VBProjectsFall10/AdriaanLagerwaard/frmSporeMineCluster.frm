VERSION 5.00
Begin VB.Form frmSporeMineCluster 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Spore Mine Cluster"
   ClientHeight    =   12345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19755
   LinkTopic       =   "Form1"
   ScaleHeight     =   12345
   ScaleWidth      =   19755
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   3615
      Left            =   7440
      ScaleHeight     =   3555
      ScaleWidth      =   4875
      TabIndex        =   8
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdSporeTotal 
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
      Left            =   4680
      TabIndex        =   7
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdSporeTotalCls 
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
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdSporeCls 
      Caption         =   "Clear Number of Spore Mines"
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
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtSpore 
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
      Left            =   3480
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picSpore 
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
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin VB.CommandButton cmdSporeMindClusterBack 
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
      Height          =   1095
      Left            =   3480
      TabIndex        =   0
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "3-6 per Brood: 10 pts/Each"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Spore Mine Cluster"
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
      Width           =   3255
   End
End
Attribute VB_Name = "frmSporeMineCluster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdSporeCls_Click()
txtSpore = 0

End Sub

Private Sub cmdSporeMindClusterBack_Click()
frmSporeMineCluster.Hide
frmFastAttack.Show
picSpore.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "SporeMineCluster" Then
            picSpore.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picSpore.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picSpore.Print "Sorry."
    End If

    
End Sub

Private Sub cmdSporeTotal_Click()
NumberSpore = txtSpore
SporeTotal = SporeTotal + (10 * NumberSpore)

MsgBox "Your Total Points Spent on Spore Mines is " & SporeTotal

End Sub

Private Sub cmdSporeTotalCls_Click()
SporeTotal = 0
txtSpore = 0

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\spore.JPG")

End Sub
