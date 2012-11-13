VERSION 5.00
Begin VB.Form frmBigLittle 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form2"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   FillColor       =   &H00C0C0FF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8775
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "quit "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton cmdLittle 
      BackColor       =   &H0000C000&
      Caption         =   "Little Sister!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   2535
   End
   Begin VB.CommandButton cmdBig 
      BackColor       =   &H0000C000&
      Caption         =   "Big Sister!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Image imgRush07 
      Height          =   5745
      Left            =   600
      Picture         =   "frmBigLittle.frx":0000
      Top             =   1200
      Width           =   9060
   End
   Begin VB.Label lblBigLittle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pick your current status!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmBigLittle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBig_Click()
    frmBigService.Show
    frmBigLittle.Hide
End Sub

Private Sub cmdLittle_Click()
    frmLittleService.Show
    frmBigLittle.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
