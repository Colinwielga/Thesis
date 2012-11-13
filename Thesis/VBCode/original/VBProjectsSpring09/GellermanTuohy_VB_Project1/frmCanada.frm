VERSION 5.00
Begin VB.Form frmCanada 
   BackColor       =   &H00004000&
   Caption         =   "Canada"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Fantastic Travel Experience With Johnnie Travel!!! :'("
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   10
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdNextActivity 
      Caption         =   "Time For You To Check Out And Get On Your Way For Your Dream Vacation With Johnnie Travel!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   6840
      TabIndex        =   9
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton cmdDeerhunt 
      Caption         =   "Add"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ComboBox cboDeerhunt 
      Height          =   315
      ItemData        =   "frmCanada.frx":0000
      Left            =   3000
      List            =   "frmCanada.frx":0022
      TabIndex        =   5
      Text            =   "(People)"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdHorseriding 
      Caption         =   "Add"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox cboHorseriding 
      Height          =   315
      ItemData        =   "frmCanada.frx":0045
      Left            =   3000
      List            =   "frmCanada.frx":0067
      TabIndex        =   2
      Text            =   "(people)"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   2280
      Picture         =   "frmCanada.frx":008A
      Top             =   4200
      Width           =   3750
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Caption         =   "$34"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      Caption         =   "$27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "Mountain Deer Hunting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800080&
      Caption         =   "Horseback Riding"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Activities in Saskatchewan, Canada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "frmCanada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Vacation Planner
'Form Name: Destination
'Authors: Luke Gellerman and Tan Tuohy
'3/22/09
'This form allows the user to sign up for activites that are particular to this Location
'If the user's Location would have been different that theychose earlier in the program, then
'they would not have been brought to this page. The activities that the user signs up for are added to the
'CheckoutTotal that is used in the last form.

Option Explicit
Private Sub cmdDeerhunt_Click()

    'declare variable
    Dim Deerhunt As Single

    'get the cost
    Deerhunt = (CInt(cboDeerhunt.Text) * 34)

    'running activity total and checkout total
    ActivitiesTotal = ActivitiesTotal + Deerhunt
    CheckoutTotal = CheckoutTotal + ActivitiesTotal

End Sub

Private Sub cmdHorseriding_Click()

    'declare variable
    Dim Horseriding As Single

    'get the cost
    Horseriding = (CInt(cboHorseriding.Text) * 27)

    'running activity total and checkout total
    ActivitiesTotal = ActivitiesTotal + Horseriding
    CheckoutTotal = CheckoutTotal + ActivitiesTotal

End Sub

Private Sub cmdNextActivity_Click()
    'hides current form and brings up checkout form
    
    frmCanada.Hide
    frmCheckout.Show
    
End Sub

Private Sub cmdQuit_Click()
    End     'ends program
End Sub

Private Sub Form_Load()
    'This code centers the form on computer screen upon loading

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
