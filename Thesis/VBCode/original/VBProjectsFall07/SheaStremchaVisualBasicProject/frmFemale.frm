VERSION 5.00
Begin VB.Form frmFemale 
   BackColor       =   &H00404040&
   Caption         =   "Female"
   ClientHeight    =   9270
   ClientLeft      =   1320
   ClientTop       =   1095
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   13095
   Begin VB.CommandButton cmdBrunette 
      BackColor       =   &H00004080&
      Caption         =   "Brunette"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdBlonde 
      BackColor       =   &H0000FFFF&
      Caption         =   "Blonde"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdRed 
      BackColor       =   &H000000FF&
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Now Select Your Hair Color"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label lblFemale 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "frmFemale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBlonde_Click()
'this sub loads the form of a blonde woman
'It also adds an "L" to the PicName variable to show the picture output will be a blond female
frmFemale.Hide
frmFBlonde.Show
PicName = PicName + "L"
End Sub

Private Sub cmdBrunette_Click()
'this sub loads the form for a brunette woman
'It also adds an "B" to the PicName variable to show the picture output will be a brown haired female
frmFemale.Hide
frmFBrown.Show
PicName = PicName + "B"
End Sub

Private Sub cmdRed_Click()
'this sub loads the form for a Red Headed woman
'It also adds an "L" to the PicName variable to show the picture output will be a red-headed female
frmFemale.Hide
frmFRed.Show
PicName = PicName + "R"
End Sub

