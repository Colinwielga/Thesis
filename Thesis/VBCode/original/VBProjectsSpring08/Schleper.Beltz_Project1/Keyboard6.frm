VERSION 5.00
Begin VB.Form Keyboard6 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   Picture         =   "Keyboard6.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000000FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Back to Table"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   7440
      Width           =   2415
   End
   Begin VB.TextBox txtMessage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   27
      Top             =   3720
      Width           =   6855
   End
   Begin VB.CommandButton CMDSPACE 
      Caption         =   "SPACE"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   26
      Top             =   6720
      Width           =   3015
   End
   Begin VB.CommandButton CMDZ 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   1800
      TabIndex        =   25
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton CMDX 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   24
      Left            =   2760
      TabIndex        =   24
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton CMDC 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   3720
      TabIndex        =   23
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton CMDV 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   4680
      TabIndex        =   22
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton CMDB 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   21
      Left            =   5640
      TabIndex        =   21
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton CMDN 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   6600
      TabIndex        =   20
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton CMDM 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   19
      Left            =   7560
      TabIndex        =   19
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton CMDA 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   18
      Left            =   840
      TabIndex        =   18
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDS 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   1800
      TabIndex        =   17
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDD 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   2760
      TabIndex        =   16
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDF 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   3720
      TabIndex        =   15
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDG 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   4680
      TabIndex        =   14
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDH 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   5640
      TabIndex        =   13
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDJ 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   6600
      TabIndex        =   12
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDK 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   7560
      TabIndex        =   11
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDL 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   8520
      TabIndex        =   10
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDW 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton CMDE 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   2280
      TabIndex        =   8
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton CMDR 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   3240
      TabIndex        =   7
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton CMDT 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   4200
      TabIndex        =   6
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton CMDY 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   5160
      TabIndex        =   5
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton CMDU 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   6120
      TabIndex        =   4
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton CMDI 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   7080
      TabIndex        =   3
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton CMDO 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   8040
      TabIndex        =   2
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton CMDP 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   9000
      TabIndex        =   1
      Top             =   4320
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2295
      Left            =   2400
      Picture         =   "Keyboard6.frx":1619F
      ScaleHeight     =   2235
      ScaleWidth      =   6045
      TabIndex        =   0
      Top             =   0
      Width           =   6105
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Message"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   30
      Top             =   3000
      Width           =   2655
   End
End
Attribute VB_Name = "Keyboard6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vinnie Joe's Pub
'Keyboard6
'Vinnie Schleper, Joey Beltz
'3/26/08
' this form is used to enter specific messages needed in the table picResults box.
Option Explicit
Private OldX As Integer
  Private OldY As Integer
  Private DragMode As Boolean
  Dim MoveMe As Boolean

  Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     MoveMe = True
     OldX = X
     OldY = Y

 End Sub

 Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


     If MoveMe = True Then
         Me.Left = Me.Left + (X - OldX)
         Me.Top = Me.Top + (Y - OldY)
     End If

 End Sub

 Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


     Me.Left = Me.Left + (X - OldX)
     Me.Top = Me.Top + (Y - OldY)
     MoveMe = False

 End Sub


Private Sub CMDA_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDA(Index).Caption
End Sub

Private Sub cmdAdd_Click()
message6 = txtMessage.Text
Table6.Show
Keyboard6.Hide
End Sub

Private Sub CMDB_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDB(Index).Caption
End Sub

Private Sub CMDC_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDC(Index).Caption
End Sub

Private Sub cmdClear_Click()
txtMessage.Text = ""
End Sub

Private Sub CMDD_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDD(Index).Caption
End Sub

Private Sub CMDE_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDE(Index).Caption
End Sub

Private Sub CMDF_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDF(Index).Caption
End Sub

Private Sub CMDG_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDG(Index).Caption
End Sub

Private Sub CMDH_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDH(Index).Caption
End Sub

Private Sub CMDI_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDI(Index).Caption
End Sub

Private Sub CMDJ_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDJ(Index).Caption
End Sub

Private Sub CMDK_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDK(Index).Caption
End Sub

Private Sub CMDL_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDL(Index).Caption
End Sub

Private Sub CMDM_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDM(Index).Caption
End Sub

Private Sub CMDN_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDN(Index).Caption
End Sub

Private Sub CMDO_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDO(Index).Caption
End Sub

Private Sub CMDP_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDP(Index).Caption
End Sub

Private Sub cmdQ_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & cmdQ(Index).Caption
End Sub

Private Sub CMDR_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDR(Index).Caption
End Sub

Private Sub CMDS_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDS(Index).Caption
End Sub

Private Sub CMDSPACE_Click()
txtMessage.Text = txtMessage.Text & " "
End Sub

Private Sub CMDT_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDT(Index).Caption
End Sub

Private Sub CMDU_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDU(Index).Caption
End Sub

Private Sub CMDV_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDV(Index).Caption
End Sub

Private Sub CMDW_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDW(Index).Caption
End Sub

Private Sub CMDX_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDX(Index).Caption
End Sub

Private Sub CMDY_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDY(Index).Caption
End Sub

Private Sub CMDZ_Click(Index As Integer)
txtMessage.Text = txtMessage.Text & CMDZ(Index).Caption
End Sub




