VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H008080FF&
   Caption         =   "Form5"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   LinkTopic       =   "Form5"
   ScaleHeight     =   8820
   ScaleWidth      =   11055
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back One Slide"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      Caption         =   $"Tuition5.frx":0000
      Height          =   1695
      Left            =   840
      TabIndex        =   1
      Top             =   5520
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "Discreationary Note:"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   5040
      Width           =   1695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Form4.Show
Form5.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub
