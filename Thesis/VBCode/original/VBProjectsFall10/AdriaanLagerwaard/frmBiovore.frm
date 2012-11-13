VERSION 5.00
Begin VB.Form frmBiovore 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Biovore"
   ClientHeight    =   12105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19575
   LinkTopic       =   "Form4"
   ScaleHeight     =   12105
   ScaleWidth      =   19575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBiovoreTotal 
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
      Left            =   4920
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdBiovoreTotalCls 
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
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdBiovoreCls 
      Caption         =   "Clear Number of Biovores"
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
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txtBiovore 
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
      Left            =   2640
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picBiovore 
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
   Begin VB.CommandButton cmdBiovoreBack 
      Caption         =   "Back To Heavy Support"
      Height          =   1215
      Left            =   4560
      TabIndex        =   0
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Equiped With:     -Claws and Teeth    -Hardened Carapace -Spore Mine Launcher"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "1-3 per Brood:    45 pts/Each"
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
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Biovore Brood"
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
      Width           =   2415
   End
End
Attribute VB_Name = "frmBiovore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdBiovoreBack_Click()
frmBiovore.Hide
frmHeavySupport.Show
picBiovore.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Biovore" Then
            picBiovore.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picBiovore.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picBiovore.Print "Sorry."
    End If

    
End Sub


Private Sub cmdBiovoreCls_Click()
txtBiovore = 0

End Sub

Private Sub cmdBiovoreTotal_Click()
NumberBiovore = txtBiovore.Text
BiovoreTotal = BiovoereTotal + (45 * NumberBiovore)

MsgBox "Your Total Points Spent on Biovores is " & BiovoreTotal

End Sub

Private Sub cmdBiovoreTotalCls_Click()
BiovoreTotal = 0
txtBiovore = 0

End Sub
