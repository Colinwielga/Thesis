VERSION 5.00
Begin VB.Form frmVenomthrope 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Venomthrope Brood"
   ClientHeight    =   13110
   ClientLeft      =   2595
   ClientTop       =   1110
   ClientWidth     =   20370
   LinkTopic       =   "Form5"
   ScaleHeight     =   13110
   ScaleWidth      =   20370
   Begin VB.CommandButton cmdVenomthropeTotal 
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
      Left            =   3360
      TabIndex        =   8
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdVenomethropeTotalCls 
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
      Left            =   3360
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdNumberVenomthropeCls 
      Caption         =   "Clear Number of Venomthropes"
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
      Left            =   3360
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.PictureBox picVenomthrope 
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
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox txtVenomthrope 
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
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdVenomthropeBack 
      Caption         =   "Back To Elites"
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Equipted With:          -Lash Whips               -Reinforce Chitin        -Toxic Miasma"
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
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "1-9, up to 3 per Brood: 55 pts/ Each"
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
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Venomthrope Brood"
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
      Width           =   3375
   End
End
Attribute VB_Name = "frmVenomthrope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdNumberVenomthropeCls_Click()
txtVenomthrope.Text = 0
VenomthropeTotal = VenomthropeTotal - (55 * NumberVenomthrope)
End Sub

Private Sub cmdVenomethropeTotalCls_Click()
VenomthropeTotal = 0

End Sub

Private Sub cmdVenomthropeBack_Click()
frmVenomthrope.Hide
frmElites.Show
picVenomthrope.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Venomthrope" Then
            picVenomthrope.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picVenomthrope.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picVenomthrope.Print "Sorry."
    End If

    
End Sub

Private Sub cmdVenomthropeTotal_Click()
Dim NumberVenomthrope As Single

NumberVenomthrope = txtVenomthrope.Text

VenomthropeTotal = VenomthropeTotal + (55 * NumberVenomthrope)
 MsgBox "You have " & NumberVenomthrope & " worth " & VenomthropeTotal
 
End Sub
