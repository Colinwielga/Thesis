VERSION 5.00
Begin VB.Form frmHiveGuard 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hive Guard Brood"
   ClientHeight    =   11805
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   ScaleHeight     =   11805
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   4215
      Left            =   7680
      ScaleHeight     =   4155
      ScaleWidth      =   4515
      TabIndex        =   9
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton cmdHiveGuardTotal 
      Caption         =   "Total points spent"
      Height          =   975
      Left            =   3600
      TabIndex        =   8
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdHiveGuardTotalCls 
      Caption         =   "Clear points spent "
      Height          =   975
      Left            =   3600
      TabIndex        =   7
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmdTxtBoxCls 
      Caption         =   "Clear Number of Hive Guards"
      Height          =   975
      Left            =   3600
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtHiveGuard 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picHiveGuard 
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   5835
      TabIndex        =   2
      Top             =   840
      Width           =   5895
   End
   Begin VB.CommandButton cmdHiveGuardBack 
      Caption         =   "Back To Elites"
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Equiped With:     -Claws and teeth -Impaler Cannon -Hardened Carapace"
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
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "1-9 , up to 3 per Brood: 50pts/ each"
      Height          =   735
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Hive Guard Brood"
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
      Width           =   3495
   End
End
Attribute VB_Name = "frmHiveGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdHiveGuardBack_Click()
frmHiveGuard.Hide
frmElites.Show

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "HiveGuard" Then
            picHiveGuard.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picHiveGuard.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "     "; Ld(j); "      "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picHiveGuard.Print "Sorry."
    End If

    
End Sub

Private Sub cmdHiveGuardTotal_Click()
Dim NumberHiveGuard As Single
NumberHiveGuard = txtHiveGuard.Text
HiveGuardTotal = HiveGuardTotal + (50 * NumberHiveGuard)

MsgBox "You have " & NumberHiveGuard & " worth. " & HiveGuardTotal


End Sub

Private Sub cmdHiveGuardTotalCls_Click()
HiveGuardTotal = 0

End Sub

Private Sub cmdTxtBoxCls_Click()
txtHiveGuard = 0
HiveGuardTotal = HiveGuardTotal - (50 * NumberHiveGuard)

End Sub


Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\hiveguard.JPG")

End Sub
