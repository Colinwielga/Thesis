VERSION 5.00
Begin VB.Form frmCompetition 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      Caption         =   "Back to Home Page"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9000
      TabIndex        =   11
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton cmdHWT 
      Caption         =   "HWT"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmd197 
      Caption         =   "197"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmd184 
      Caption         =   "184"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmd174 
      Caption         =   "174"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmd165 
      Caption         =   "165"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmd157 
      Caption         =   "157"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmd149 
      Caption         =   "149"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmd141 
      Caption         =   "141"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmd133 
      Caption         =   "133"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmd125 
      Caption         =   "125"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a Weight Class"
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
      Left            =   3600
      TabIndex        =   12
      Top             =   0
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   9180
      Left            =   0
      Picture         =   "frmCompetition.frx":0000
      Top             =   0
      Width           =   12240
   End
End
Attribute VB_Name = "frmCompetition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Weights(1 To 15) As String, CTR2 As Integer, Class As Integer
'shows which class you will be going up against depending on your weight class you select

Private Sub cmd125_Click()
frm125.Show 'Show the 125 weight class
frmCompetition.Hide 'hide the Competition form

End Sub

Private Sub cmd133_Click()
frm133.Show 'shows the 133 weight class
frmCompetition.Hide 'hides the Competition form
End Sub

Private Sub cmd141_Click()
frm141.Show 'shows the 141 weight class
frmCompetition.Hide 'hides the Competition form

End Sub

Private Sub cmd149_Click()
    frm149.Show 'shows the 149 weight class
    frmCompetition.Hide 'hides the Competition form
    
End Sub

Private Sub cmd157_Click()
frm157.Show 'shows the 157 weight class
frmCompetition.Hide 'hides the Competition form
End Sub
Private Sub cmd165_Click()
frm165.Show 'shows the 165 weight class
frmCompetition.Hide 'hides the Competition form
End Sub

Private Sub cmd174_Click()
frm174.Show 'shows the 174 weight class
frmCompetition.Hide 'hides the Competition form
End Sub

Private Sub cmd184_Click()
frm184.Show 'shows the 184 weight class
frmCompetition.Hide 'hides the Competition form
End Sub

Private Sub cmd197_Click()
frm197.Show 'shows the 197 weight class
frmCompetition.Hide 'hides the Competition form
End Sub

Private Sub cmdHome_Click()
frmCompetition.Hide 'hides the Competition form
frmHome.Show 'brings the viewer back to the home page
End Sub

Private Sub cmdHWT_Click()
frmHWT.Show 'shows the Heavyweight weight class
frmCompetition.Hide 'hides the Competition form
End Sub

Private Sub cmdQuit_Click()
    End 'ends program
End Sub
