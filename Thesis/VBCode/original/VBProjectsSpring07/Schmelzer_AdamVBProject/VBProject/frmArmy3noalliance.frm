VERSION 5.00
Begin VB.Form frmArmy3noalliance 
   BackColor       =   &H00004040&
   Caption         =   "The Army"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   Picture         =   "frmArmy3noalliance.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKnights1 
      Caption         =   "There are 715 battle-hardened Squires ready to be promoted to Knights at your calling."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   14
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdKnights3 
      Caption         =   "There are 555 battle-hardened Squires ready to be promoted to Knights at your calling."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   13
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdKnights2 
      Caption         =   "There are 630 battle-hardened Squires ready to be promoted to Knights at your calling."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdArchers 
      Caption         =   "Ready 2,125 Archers"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   11
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdSiege 
      Caption         =   "Ready 50 Siege Engines"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   10
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdPikemen1 
      Caption         =   "Ready 7,250 Pikemen"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   9
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdPikemen2 
      Caption         =   "Ready 6,250 Pikemen"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   8
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdPikemen3 
      Caption         =   "Ready 5,500 Pikemen"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   7
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdArchers1 
      Caption         =   "Ready 1,815 Archers"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdArchers2 
      Caption         =   "Ready 1,562 Archers"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdArchers3 
      Caption         =   "Ready 1,375  Archers"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdKnights 
      Caption         =   "There are 825 battle-hardened Squires ready to be promoted to Knights at your calling."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   3
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdPikemen 
      Caption         =   "Ready 8,500 Pikemen"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   1
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00004040&
      Caption         =   $"frmArmy3noalliance.frx":AAD7
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   8655
   End
End
Attribute VB_Name = "frmArmy3noalliance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'depending on the resource level the user has accumulated or lost, different levels of
'troops available to create are made visible as command buttons
'the code is the same for every command button option, excpet how it affects the
'battlepoints variable
'the code brings the user to either the 1on1 form, the assassination form, or
'the final battle form depending on what his situation is concerning alliances
'and the outcome of the first seige in frmArmy2

Private Sub cmdArchers_Click()
archers = archers + 2125
Battlepoints = Battlepoints + (2125 * 4)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
End Sub

Private Sub cmdArchers1_Click()
Battlepoints = Battlepoints + (1815 * 4)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdArchers2_Click()
Battlepoints = Battlepoints + (1562 * 4)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdArchers3_Click()
Battlepoints = Battlepoints + (1375 * 4)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdKnights1_Click()
Battlepoints = Battlepoints + (715 * 10)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdKnights_Click()
knights = knights + 825
Battlepoints = Battlepoints + (825 * 10)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
End Sub


Private Sub cmdKnights2_Click()
Battlepoints = Battlepoints + (630 * 10)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdKnights3_Click()
Battlepoints = Battlepoints + (555 * 10)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdPikemen_Click()
pikemen = pikemen + 8500
Battlepoints = Battlepoints + (8500 * 1)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
End Sub

Private Sub cmdPikemen1_Click()
Battlepoints = Battlepoints + 7250
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdPikemen2_Click()
Battlepoints = Battlepoints + 6250
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdPikemen3_Click()
Battlepoints = Battlepoints + 5500
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub


Private Sub cmdSengines_Click()
'no points since there will be no siege
siege = siege + 50
Battlepoints = Battlepoints + 0
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSiege_Click()
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3noalliance.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3noalliance.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3noalliance.Hide
    frmFinalbattle.Show
End If
End Sub

Private Sub Form_Load()
If Resources > 850 Then
    cmdArchers.Visible = True
    cmdArchers1.Visible = False
    cmdArchers2.Visible = False
    cmdArchers3.Visible = False
    cmdKnights.Visible = True
    cmdKnights1.Visible = False
    cmdKnights2.Visible = False
    cmdKnights3.Visible = False
    cmdPikemen.Visible = True
    cmdPikemen1.Visible = False
    cmdPikemen2.Visible = False
    cmdPikemen3.Visible = False
End If
If Resources > 500 Then
    cmdArchers.Visible = False
    cmdArchers1.Visible = True
    cmdArchers2.Visible = False
    cmdArchers3.Visible = False
    cmdKnights.Visible = False
    cmdKnights1.Visible = True
    cmdKnights2.Visible = False
    cmdKnights3.Visible = False
    cmdPikemen.Visible = False
    cmdPikemen1.Visible = True
    cmdPikemen2.Visible = False
    cmdPikemen3.Visible = False
End If
If Resources > 0 Then
    cmdArchers.Visible = False
    cmdArchers1.Visible = False
    cmdArchers2.Visible = True
    cmdArchers3.Visible = False
    cmdKnights.Visible = False
    cmdKnights1.Visible = False
    cmdKnights2.Visible = True
    cmdKnights3.Visible = False
    cmdPikemen.Visible = False
    cmdPikemen1.Visible = False
    cmdPikemen2.Visible = True
    cmdPikemen3.Visible = False
Else
    cmdArchers.Visible = False
    cmdArchers1.Visible = False
    cmdArchers2.Visible = False
    cmdArchers3.Visible = True
    cmdKnights.Visible = False
    cmdKnights1.Visible = False
    cmdKnights2.Visible = False
    cmdKnights3.Visible = True
    cmdPikemen.Visible = False
    cmdPikemen1.Visible = False
    cmdPikemen2.Visible = False
    cmdPikemen3.Visible = True
End If
End Sub
