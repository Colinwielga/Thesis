VERSION 5.00
Begin VB.Form frmOldOneEye 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Old One Eye"
   ClientHeight    =   12180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19605
   LinkTopic       =   "Form5"
   ScaleHeight     =   12180
   ScaleWidth      =   19605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOldTotal 
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
      TabIndex        =   8
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdOldTotalCls 
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
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdOldCls 
      Caption         =   "Clear Number of Old One Eyes"
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
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
   Begin VB.PictureBox picOld 
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
      TabIndex        =   4
      Top             =   840
      Width           =   5295
   End
   Begin VB.TextBox txtOld 
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
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdOldOneEyeBack 
      Caption         =   "Back To Heavy Support"
      Height          =   1095
      Left            =   5040
      TabIndex        =   0
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Equiped With:        -Bonded Exoskeleton  -Crushing Claws            -Scything Talons"
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
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "1 per Brood:      260 pts/Each"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Old One Eye"
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
      Width           =   2295
   End
End
Attribute VB_Name = "frmOldOneEye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdOldCls_Click()
txtOld = 0

End Sub

Private Sub cmdOldOneEyeBack_Click()
frmOldOneEye.Hide
frmHeavySupport.Show
picOld.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "OldOneEye" Then
            picOld.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picOld.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picOld.Print "Sorry."
    End If

    
End Sub


Private Sub cmdOldTotal_Click()
NumberOld = txtOld
OldTotal = OldTotal + (260 * NumberOld)

MsgBox "Your Total Points Spent on Old One Eye is " & OldTotal

End Sub

Private Sub cmdOldTotalCls_Click()
OldTotal = 0
txtOld = 0

End Sub
