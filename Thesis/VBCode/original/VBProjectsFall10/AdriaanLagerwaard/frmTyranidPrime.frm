VERSION 5.00
Begin VB.Form frmPrime 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tyranid Prime"
   ClientHeight    =   11940
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   11940
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrimeTotal 
      Caption         =   "Total Points spent"
      Height          =   1215
      Left            =   7800
      TabIndex        =   20
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton cmdPrimeCheckBoxCls 
      Caption         =   "Clear Check Boxes"
      Height          =   1095
      Left            =   4320
      TabIndex        =   19
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrimeTotalCls 
      Caption         =   "Clear Total Points "
      Height          =   1095
      Left            =   600
      TabIndex        =   18
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CheckBox Check11 
      Caption         =   "80 pts"
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   120
      Width           =   3015
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Regeneration 10pts"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Toxin Sacs 10pts"
      Height          =   495
      Left            =   4320
      TabIndex        =   14
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Adrenal Glands 10pts"
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CheckBox Check7 
      Caption         =   "An additional set of scything talons Free"
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   5160
      Width           =   3375
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Deathspitter 5pts"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   4800
      Width           =   3375
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Spinefists Free"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Rending Claws Free"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Width           =   3375
   End
   Begin VB.CheckBox Check3 
      Caption         =   "LashWhip and BoneSword 15pts"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "A pair of BoneSwords 10pts"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Rending Claws 5pts"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   3375
   End
   Begin VB.PictureBox picPrime 
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
      Left            =   360
      ScaleHeight     =   915
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   840
      Width           =   7335
   End
   Begin VB.CommandButton cmdTyrandiPrimeBack 
      Caption         =   "Go back to HQ choices"
      Height          =   1215
      Left            =   7920
      TabIndex        =   0
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Equiped with:                -Bonded Exoskeleton  -Devourer                       -Scything Talons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7920
      TabIndex        =   17
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "Take any of the following:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Replace Devourer with:"
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
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Replace scything talons with:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label 
      Caption         =   "Tyranid Prime"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmPrime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms
'The checkBoxes will add points from values stated in the module to the Equipment cost
'that is later added to the Total Value

Private Sub Check1_Click()
PrimeTotal = PrimeTotal + RendingClaws

End Sub

Private Sub Check10_Click()
PrimeTotal = PrimeTotal + TPRegeneration

End Sub

Private Sub Check11_Click()
PrimeTotal = PrimeTotal + 80

End Sub

Private Sub Check2_Click()
PrimeTotal = PrimeTotal + BoneSword

End Sub

Private Sub Check3_Click()
PrimeTotal = PrimeTotal + BoneSword + LashWhip

End Sub

Private Sub Check6_Click()
PrimeTotal = PrimeTotal + Deathspitter
End Sub

Private Sub Check8_Click()
PrimeTotal = PrimeTotal + AdrenalGlands

End Sub

Private Sub Check9_Click()
PrimeTotal = PrimeTotal + ToxinSacs

End Sub

Private Sub cmdPrimeCheckBoxCls_Click()
Check1 = 0
Check2 = 0
Check3 = 0
Check4 = 0
Check5 = 0
Check6 = 0
Check7 = 0
Check8 = 0
Check9 = 0
Check10 = 0
Check11 = 0
PrimeTotal = PrimeTotal / 2

End Sub

Private Sub cmdPrimeTotal_Click()
MsgBox "This is how much you have spent on a Tyranid Prime. " & PrimeTotal & "pts."

End Sub

Private Sub cmdPrimeTotalCls_Click()
PrimeTotal = 0

End Sub

Private Sub cmdTyrandiPrimeBack_Click()
frmPrime.Hide
frmHQ.Show
picPrime.Cls

End Sub

Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Prime" Then
            picPrime.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picPrime.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picPrime.Print "Sorry."
    End If

    
End Sub

