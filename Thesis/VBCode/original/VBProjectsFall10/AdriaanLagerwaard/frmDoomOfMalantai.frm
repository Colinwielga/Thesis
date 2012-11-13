VERSION 5.00
Begin VB.Form frmDoomOfMalantai 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Doom of Malant'tai"
   ClientHeight    =   11355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17310
   LinkTopic       =   "Form3"
   ScaleHeight     =   11355
   ScaleWidth      =   17310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDoomOfMalantaiTotal 
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
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdDoomOfMalantaiTotalCls 
      Caption         =   "Clear Pionts Spent"
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
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdDoomOfMalantaiCheckBoxCls 
      Caption         =   "Clear Check Box"
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
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "90 pts"
      Height          =   615
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox picDoomOfMalantai 
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
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   840
      Width           =   6255
   End
   Begin VB.CommandButton cmdDoomOfMalantaiBack 
      Caption         =   "Back To Elites"
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
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Equiped With:      -Claws and Teeth -Reinforce Chitin  Psychic Powers:  -Cataclysm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "The Doom Of Malant'tai"
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmDoomOfMalantai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub Check1_Click()
DoomOfMalantaiTotal = DoomOfMalantaiTotal + 90

End Sub

Private Sub cmdDoomOfMalantaiBack_Click()
frmDoomOfMalantai.Hide
frmElites.Show
picDoomOfMalantai.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "DoomOfMalantai" Then
            picDoomOfMalantai.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picDoomOfMalantai.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picDoomOfMalantai.Print "Sorry."
    End If
End Sub

Private Sub cmdDoomOfMalantaiCheckBoxCls_Click()
Check1 = 0
DoomOfMalantaiTotal = DoomOfMalantaiTotal / 2


End Sub

Private Sub cmdDoomOfMalantaiTotal_Click()
MsgBox "Total Points Spent " & DoomOfMalantaiTotal

End Sub

Private Sub cmdDoomOfMalantaiTotalCls_Click()
DoomOfMalantaiTotal = 0

End Sub

