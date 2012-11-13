VERSION 5.00
Begin VB.Form frmArtHistoryOpen 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Renaissance to Expressionism"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6195
   DrawWidth       =   5
   FillColor       =   &H00000080&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MouseIcon       =   "project1.from.frx":0000
   ScaleHeight     =   2865
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00AC796C&
      Caption         =   "Exit Program"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetStarted 
      Appearance      =   0  'Flat
      BackColor       =   &H00AC796C&
      Caption         =   "Get Started"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      MousePointer    =   4  'Icon
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      Picture         =   "project1.from.frx":030A
      ScaleHeight     =   975
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmArtHistoryOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetStarted_Click()
    frmArtHistoryOpen.Visible = False
    frmUsr_Info.Visible = True
    
End Sub

Private Sub Command1_Click()
    End
End Sub
