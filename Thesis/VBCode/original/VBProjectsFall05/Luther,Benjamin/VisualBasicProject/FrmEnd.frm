VERSION 5.00
Begin VB.Form FrmEnd 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Exit?"
   ClientHeight    =   1770
   ClientLeft      =   6690
   ClientTop       =   4590
   ClientWidth     =   2535
   LinkTopic       =   "Form2"
   ScaleHeight     =   1770
   ScaleWidth      =   2535
   Begin VB.CommandButton Cmdnoend 
      BackColor       =   &H00FFC0C0&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Cmdend 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Are you sure you want to exit?"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FrmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SurfProject (SurfingProject.vbp)
'Form Name: frmEnd (frmEnd.frm)
'Author: Benjamin Luther
'Purpose of Form: 'This form is designed to assure that
                        'the user is ready to quit the program, it asks
                        'them if they are sure they are ready to exit.
Private Sub Cmdend_Click()
    End 'ends the program
End Sub

Private Sub Cmdnoend_Click()
    FrmEnd.Hide 'hides the end form, returns to the main page
End Sub
