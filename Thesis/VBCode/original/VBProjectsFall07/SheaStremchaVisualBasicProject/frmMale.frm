VERSION 5.00
Begin VB.Form frmMale 
   BackColor       =   &H00404040&
   Caption         =   "Male"
   ClientHeight    =   9240
   ClientLeft      =   1320
   ClientTop       =   1305
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   13170
   Begin VB.CommandButton cmdBrunette 
      BackColor       =   &H00004080&
      Caption         =   "Brunette"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdBlonde 
      BackColor       =   &H0000FFFF&
      Caption         =   "Blonde"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdRed 
      BackColor       =   &H000000FF&
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblHair 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Now Select Your Hair Color"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmMale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBlonde_Click()
'This Sub loads the form for a Blonde Man
'Adds an "L" to the PicName variable for Blonde
frmMale.Hide
frmMBlonde.Show
PicName = PicName + "L"
End Sub

Private Sub cmdBrunette_Click()
'This sub loads the form for a brunette man
'Adds an "B" to the PicName variable for Brunette
frmMale.Hide
frmMBrown.Show
PicName = PicName + "B"
End Sub

Private Sub cmdRed_Click()
'this sub loads the form for a red headed man
'Adds an "R" to the PicName variable for Red Head
frmMale.Hide
frmMRed.Show
PicName = PicName + "R"
End Sub
