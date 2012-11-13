VERSION 5.00
Begin VB.Form frmYmgarlGenestealers 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ymgarl Genestealers Brood"
   ClientHeight    =   12120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19380
   LinkTopic       =   "Form1"
   ScaleHeight     =   12120
   ScaleWidth      =   19380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdYmgarlGenestealersTotal 
      Caption         =   "Total Pionts Spent"
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
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdYmgarlGenestealersTotalCls 
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
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdYmgarlGenestealersCls 
      Caption         =   "Clear Number of Ymgarl Genestealers"
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
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtYmgarlGenestealers 
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
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox picYmgarlGenestealers 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   1080
      Width           =   5655
   End
   Begin VB.CommandButton cmdYmgarlGenestealersBack 
      Caption         =   "Back To Elites"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Equiped With:           -Hardened Carapace  -Rending Claws"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "5-30, 5-10 per Brood: 23 pts/Each"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Ymgarl Genestealers Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmYmgarlGenestealers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub cmdYmgarlGenestealersCls_Click()
txtYmgarlGenestealers.Text = 0
YmgarlGenestealersTotal = YmgarlGenestealersTotal - (23 * NumberYmgarlGenestealers)
End Sub
Private Sub cmdYmgarlGenestealersBack_Click()
frmYmgarlGenestealers.Hide
frmElites.Show
picYmgarlGenestealers.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "YmgarlGenestealers" Then
            picYmgarlGenestealers.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picYmgarlGenestealers.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picYmgarlGenestealers.Print "Sorry."
    End If

    
End Sub


Private Sub cmdYmgarlGenestealersTotal_Click()
Dim NumberYmgarlGenestealers As Single
NumberYmgarlGenestealers = txtYmgarlGenestealers.Text
YmgarlGenestealersTotal = YmgarlGenestealersTotal + (23 * NumberYmgarlGenestealers)

MsgBox "You Have " & NumberYmgarlGenestealers & "worth" & YmgarlGenestealersTotal

End Sub

Private Sub cmdYmgarlGenestealersTotalCls_Click()
YmgarlGenestealersTotal = 0

End Sub
