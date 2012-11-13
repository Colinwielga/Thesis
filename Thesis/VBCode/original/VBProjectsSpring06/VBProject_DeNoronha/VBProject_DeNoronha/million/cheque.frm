VERSION 5.00
Begin VB.Form frmCheque 
   BackColor       =   &H00000000&
   Caption         =   "Cheque"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form3"
   ScaleHeight     =   5190
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00000000&
      Caption         =   "Exit Program"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   12
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtAmt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   8160
      TabIndex        =   11
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   540
      Left            =   2280
      TabIndex        =   10
      Top             =   2355
      Width           =   4935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   840
      Picture         =   "cheque.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblBank 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "World Bank"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   3000
      TabIndex        =   9
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label lblRouting 
      BackColor       =   &H00000000&
      Caption         =   ":5845312000:-   405035458:-   2300"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      X1              =   840
      X2              =   4320
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblMeno 
      BackColor       =   &H00000000&
      Caption         =   "Memo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   2760
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   960
      X2              =   7800
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblOrder 
      BackColor       =   &H00000000&
      Caption         =   "Pay to the order of "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblNo 
      BackColor       =   &H00000000&
      Caption         =   "No. 2300"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   8280
      X2              =   9600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00000000&
      Caption         =   "March 22nd, 2006"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblCheque 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   9975
   End
   Begin VB.Label lblDesigner 
      BackColor       =   &H00000000&
      Caption         =   "Pradeep de Noronha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
End
Attribute VB_Name = "frmcheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Who wants to be a Millionare.(millionare1.vbp)

'Form name: frmCheque; Form caption: Cheque

'Author: Pradeep de Noronha

'Date written: 15th March, 2006

'Form Objective: This form is designed to display the users winnings. The users
'                winnings and name is displayed on the cheque. The commond button
'                allows the user to exit the program.

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Load()
    txtName.Text = username
    txtAmt.Text = amount
End Sub
