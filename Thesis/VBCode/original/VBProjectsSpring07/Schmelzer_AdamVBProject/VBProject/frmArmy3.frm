VERSION 5.00
Begin VB.Form frmArmy3 
   BackColor       =   &H00004040&
   Caption         =   "The Army"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   Picture         =   "frmArmy3.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPikemen2 
      Caption         =   "Ready 3,750 Pikemen"
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
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdPikemen1 
      Caption         =   "Ready 4,750 Pikemen"
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
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdArchers3 
      Caption         =   "Ready 750 Archers"
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
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdArchers2 
      Caption         =   "Ready 950 Archers"
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
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdPikemen3 
      Caption         =   "Ready 3,000 Pikemen"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdArchers1 
      Caption         =   "Ready 1,100 Archers"
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
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdKnights1 
      Caption         =   "There are 475 battle-hardened Squires ready to be promoted to Knights at your calling."
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
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdKnights2 
      Caption         =   "There are 375 battle-hardened Squires ready to be promoted to Knights at your calling."
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
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdKnights3 
      Caption         =   "There are 300 battle-hardened Squires ready to be promoted to Knights at your calling."
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
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdKnights 
      Caption         =   "There are 500 battle-hardened Squires ready to be promoted to Knights at your calling."
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
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdSengines 
      Caption         =   "Build 50 Siege Engines"
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
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdArchers 
      Caption         =   "Ready  1,500 Archers"
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
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdPikemen 
      Caption         =   "Ready 6,000 pikemen"
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
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
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
      Height          =   1575
      Left            =   8760
      TabIndex        =   1
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00004040&
      Caption         =   $"frmArmy3.frx":AAD7
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
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   8655
   End
End
Attribute VB_Name = "frmArmy3"
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
archers = archers + 1500
Battlepoints = Battlepoints + (1500 * 4)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
End Sub

Private Sub cmdArchers1_Click()
Battlepoints = Battlepoints + (1100 * 4)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdArchers2_Click()
Battlepoints = Battlepoints + (950 * 4)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdArchers3_Click()
Battlepoints = Battlepoints + (750 * 4)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdKnights1_Click()
Battlepoints = Battlepoints + (475 * 10)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdKnights_Click()
knights = knights + 50
Battlepoints = Battlepoints + (500 * 10)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
End Sub


Private Sub cmdKnights2_Click()
Battlepoints = Battlepoints + (375 * 10)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdKnights3_Click()
Battlepoints = Battlepoints + (300 * 10)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdPikemen_Click()
pikemen = pikemen + 6000
Battlepoints = Battlepoints + (6000 * 1)
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
End Sub

Private Sub cmdPikemen1_Click()
Battlepoints = Battlepoints + 4750
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdPikemen2_Click()
Battlepoints = Battlepoints + 3750
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'goes to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdPikemen3_Click()
Battlepoints = Battlepoints + 3000
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'goes to the option of killing Lannister in the final battle, which has essentially already been won
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSengines_Click()
'no points since there will be no siege
siege = siege + 50
Battlepoints = Battlepoints + 0
If failsiegeV = True Or (blockadeV = True And LannisterAllianceN = False And LannisterAllianceP = False) Then
    frmArmy3.Hide
    frm1on1.Show
End If
If LannisterAllianceN = True Then
    frmArmy3.Hide
    frmAssassination.Show
End If
If LannisterAllianceP = True Or (LannisterAllianceP = False And waitedV = True) Or (successfulsiegeV = True And LannisterAllianceP = False And LannisterAllianceN = False) Then
    frmArmy3.Hide
    frmFinalbattle.Show
End If
    'go to the option of killing Lannister in the final battle, which has essentially already been won
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
