VERSION 5.00
Begin VB.Form frmjoetown 
   Caption         =   "Tour De St. Joe"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   14160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      BackColor       =   &H00808080&
      Caption         =   "End your Tour De St. Joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdboob 
      BackColor       =   &H00C0C0FF&
      Caption         =   "The Boobery"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdtellies 
      BackColor       =   &H001E8E4A&
      Caption         =   "Tellies"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdpolice 
      BackColor       =   &H00B75C3E&
      Caption         =   "Police"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdsals 
      BackColor       =   &H00375C7D&
      Caption         =   "Sal's "
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdpaso 
      BackColor       =   &H0003CCE9&
      Caption         =   "El Paso"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lbldumb 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   11640
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lbldirection 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on where you would like to go"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   600
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   8640
      Left            =   0
      Picture         =   "frmjoetown.frx":0000
      Top             =   600
      Width           =   13050
   End
End
Attribute VB_Name = "frmjoetown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 
    'Project name:  Tour De St. Joe
    'Form:  frmjoetown, "Tour de St. Joe"
    'Author:  Brooke and Josh
    'Date:  3/11/08
    'Objective: The main page in which a user can select which places they would like to explore.  This is the main page.


Private Sub cmdboob_Click()

    frmboob.Show
    frmjoetown.Hide
    

End Sub

Private Sub cmdgo_Click()
    End
End Sub

Private Sub cmdpaso_Click()
    
    frmpaso.Show           'shows the main form
    frmjoetown.Hide        'hides the matching form
    
End Sub

Private Sub cmdpolice_Click()

    frmpolice.Show
    frmjoetown.Hide

End Sub

Private Sub cmdsals_Click()

    frmsals.Show
    frmjoetown.Hide

End Sub

Private Sub cmdtellies_Click()

    frmtellies.Show
    frmjoetown.Hide
    
End Sub
