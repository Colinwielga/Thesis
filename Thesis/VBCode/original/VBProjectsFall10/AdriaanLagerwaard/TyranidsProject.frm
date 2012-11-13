VERSION 5.00
Begin VB.Form frmParasiteOfMortrex 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Parasite of Mortrex"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total Points Spent"
      Height          =   975
      Left            =   3600
      TabIndex        =   7
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdTotalCls 
      Caption         =   "Clear Parasite of Mortrex Total"
      Height          =   975
      Left            =   3600
      TabIndex        =   6
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmdClearChecks 
      Caption         =   "Clear CheckBox"
      Height          =   975
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "160 pts"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox PicParasiteOfMortrex 
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
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   840
      Width           =   6135
   End
   Begin VB.CommandButton cmdParasiteOfMortrexBack 
      Caption         =   "Go back to HQ choices"
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   " Equiped With             -Implant Attack            -Bonded Exoskeleton -Rending Claws          -Wings"
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
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "The Parasite of Mortrex"
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
      Width           =   3855
   End
End
Attribute VB_Name = "frmParasiteOfMortrex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub Check1_Click()
ParasiteOfMortrexTotal = ParasiteOfMortrexTotal + 160

End Sub

Private Sub cmdClearChecks_Click()
Check1 = 0
ParasiteOfMortrexTotal = ParasiteOfMortrexTotal / 2
End Sub

Private Sub cmdParasiteOfMortrexBack_Click()
frmParasiteOfMortrex.Hide
frmHQ.Show

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "ParasiteOfMortrex" Then
            PicParasiteOfMortrex.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            PicParasiteOfMortrex.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        PicParasiteOfMortrex.Print "Sorry."
    End If

    
End Sub



Private Sub cmdTotalCls_Click()
ParasiteOfMortrexTotal = 0

End Sub

Private Sub cmdTotal_Click()
MsgBox "Total points spent on the Parasite of Mortrex are " & ParasiteOfMortrexTotal
End Sub
