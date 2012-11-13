VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   11805
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H008080FF&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      TabIndex        =   1
      Top             =   6000
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      Height          =   5775
      Left            =   2280
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5715
      ScaleMode       =   0  'User
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "As A CSB Resident Assistant"
         BeginProperty Font 
            Name            =   "PaintStroke"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2535
         Left            =   2640
         TabIndex        =   2
         Top             =   2400
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnter_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub Form_Load()
    strPath = "N:\CS130\handin\kmfoss\"
End Sub
