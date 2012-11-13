VERSION 5.00
Begin VB.Form frmZoanthrope 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Zoanthrope Brood"
   ClientHeight    =   12390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19665
   LinkTopic       =   "Form4"
   ScaleHeight     =   12390
   ScaleWidth      =   19665
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   4215
      Left            =   6480
      ScaleHeight     =   4155
      ScaleWidth      =   3915
      TabIndex        =   9
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton cmdZoanthropeTotal 
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
      Left            =   2880
      TabIndex        =   8
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdZoanthropeTotalCls 
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
      Left            =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmdNumberZoanthropeCls 
      Caption         =   "Clear Number of Zoanthropes"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtZoanthrope 
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
      Left            =   3840
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picZoanthrope 
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
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   840
      Width           =   5175
   End
   Begin VB.CommandButton cmdZoanthropeBack 
      Caption         =   "Back To Elites"
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Equiped With:     -Claws And Teeth -Reinforced Chitin Psychic Powers:  -Warp Blast           -Warp Lance"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "1-9, up to 3 per Brood: 60pts/Each"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Zoanthrope Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmZoanthrope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms


Private Sub cmdNumberZoanthropeCls_Click()
txtZoanthrope.Text = 0
ZoanthropeTotal = ZoanthropeTotal - (60 * NumberZoanthorpe)

End Sub

Private Sub cmdZoanthropeBack_Click()
frmZoanthrope.Hide
frmElites.Show
picZoanthrope.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Zoanthrope" Then
            picZoanthrope.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picZoanthrope.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picZoanthrope.Print "Sorry."
    End If

    
End Sub

Private Sub cmdZoanthropeTotal_Click()
Dim NumberZoanthrope As Single
NumberZoanthrope = txtZoanthrope.Text
ZoanthropeTotal = ZoanthropeTotal + (60 * NumberZoanthrope)

MsgBox "You have " & NumberZoanthrope & " worth. " & ZoanthropeTotal

End Sub

Private Sub cmdZoanthropeTotalCls_Click()
ZoanthropeTotal = 0

End Sub
Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\zoanthrope.JPG")

End Sub
