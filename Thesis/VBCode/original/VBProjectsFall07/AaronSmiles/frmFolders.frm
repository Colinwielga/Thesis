VERSION 5.00
Begin VB.Form frmFolders 
   BackColor       =   &H80000002&
   Caption         =   "Folders"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   6120
      Width           =   2055
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   2760
      Picture         =   "frmFolders.frx":0000
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   3
      Top             =   3120
      Width           =   2310
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   120
      Picture         =   "frmFolders.frx":102B
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   2
      Top             =   3120
      Width           =   2310
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   2295
      Left            =   2760
      Picture         =   "frmFolders.frx":218A
      ScaleHeight     =   2235
      ScaleWidth      =   2250
      TabIndex        =   1
      Top             =   240
      Width           =   2310
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   120
      Picture         =   "frmFolders.frx":30B2
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   240
      Width           =   2310
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "N: Drive\CS130"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "M: Drive"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "iTunes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "My Documents"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
End
Attribute VB_Name = "frmFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'these link to folders I find useful

Private Sub Command1_Click()
frmFolders.Hide
End Sub

Private Sub Picture1_Click()
    Shell ("explorer.exe \\ad\homedir$\Students\A\a1smiles\My Documents")
End Sub

Private Sub Picture2_Click()
    Shell ("explorer.exe M:\My Documents\My Music\iTunes")
End Sub

Private Sub Picture3_Click()
    Shell ("explorer.exe M:\")
End Sub

Private Sub Picture4_Click()
    Shell ("explorer.exe N:\CS130")
End Sub

