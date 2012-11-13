VERSION 5.00
Begin VB.Form Wapping 
   BackColor       =   &H008080FF&
   Caption         =   "Wapping"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   17520
      TabIndex        =   8
      Top             =   10560
      Width           =   1455
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15240
      TabIndex        =   7
      Top             =   10560
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   360
      Picture         =   "Wapping.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   9555
      TabIndex        =   3
      Top             =   1680
      Width           =   9615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   2
      Text            =   "Butlers Wharf"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   6720
      Picture         =   "Wapping.frx":3BE4
      ScaleHeight     =   2955
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   5640
      Width           =   9855
   End
   Begin VB.Label Label5 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   11520
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Two famous sites within the Wapping District of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Museum was set up by Terence Conran. and has exhibitions of 20thC design concentrating mainly on domestic objects."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   4800
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Design Museum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "You are on a river looking north towards Butlers Wharf.   Orinally warehouses - now shops, cafes and housing. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9840
      TabIndex        =   1
      Top             =   9000
      Width           =   3375
   End
End
Attribute VB_Name = "Wapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: Wapping (Wapping.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: This purpose of this form is to inform the user of the history of the Butlers Wharf and
                'the design museum
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'User returns to choose a new district on the map of London page
Wapping.Hide
MapLondon.Show
End Sub
