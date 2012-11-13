VERSION 5.00
Begin VB.Form frmHargrave 
   BackColor       =   &H80000005&
   Caption         =   "Hargrave"
   ClientHeight    =   11220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18105
   LinkTopic       =   "Form1"
   ScaleHeight     =   11220
   ScaleWidth      =   18105
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   9600
      ScaleHeight     =   1455
      ScaleWidth      =   7215
      TabIndex        =   9
      Top             =   1680
      Width           =   7215
   End
   Begin VB.CommandButton cmdWhy 
      BackColor       =   &H000080FF&
      Caption         =   "Why I Went Here?"
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdRegular 
      BackColor       =   &H000080FF&
      Caption         =   "View Regular Graduates"
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdHonor 
      BackColor       =   &H000080FF&
      Caption         =   "View Honor Graduates"
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.PictureBox picHargrave3 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   360
      ScaleHeight     =   6015
      ScaleWidth      =   6015
      TabIndex        =   4
      Top             =   240
      Width           =   6015
   End
   Begin VB.PictureBox picHargrave2 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   360
      ScaleHeight     =   4455
      ScaleWidth      =   9015
      TabIndex        =   3
      Top             =   6480
      Width           =   9015
   End
   Begin VB.PictureBox picHargrave1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   10320
      ScaleHeight     =   4935
      ScaleWidth      =   6495
      TabIndex        =   2
      Top             =   3480
      Width           =   6495
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000080FF&
      Caption         =   "Back to Main Form"
      Height          =   855
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Exit Program"
      Height          =   855
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label lblShoutOut 
      BackColor       =   &H00FF0000&
      Caption         =   "CLICK ON EITHER OF THESE BUTTONS FOR SOME COOL ARRAYS!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   6600
      TabIndex        =   11
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Tim Johnson   3/20    To State Why I went to Hargrave and to show a bit about it"
      Height          =   615
      Left            =   13680
      TabIndex        =   10
      Top             =   10440
      Width           =   2895
   End
   Begin VB.Label lblHargraveTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Hargrave Military Academy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   7320
      TabIndex        =   5
      Top             =   480
      Width           =   9255
   End
End
Attribute VB_Name = "frmHargrave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()      'Goes back to Main Form
frmMain.Show                     'Goes back to Main Form
frmHargrave.Hide
End Sub

Private Sub cmdHonor_Click()     'Goes to Honor Grads Form
frmHonorGrads.Show               'Goes to Honor Grads Form
frmHargrave.Hide
End Sub

Private Sub cmdQuit_Click()      'Ends program where you are
    End                          'Ends program where you are
End Sub

Private Sub cmdRegular_Click()   'Goes to Regular Grads Form
frmRegularGrads.Show             'Goes to Regular Grads Form
frmHargrave.Hide
End Sub

Private Sub cmdWhy_Click()       'Answers a simple question

picInfo.Print "I went to Hargrave Military Academy because my cumulative GPA after my first junior year of high"
picInfo.Print "  school was less than a 2.0 and my parents thought I needed some discipline."
picInfo.Print "  Hargrave set me on the right track and taught me the discipline to do my homework and get"
picInfo.Print "  accepted to a good college."

End Sub

Private Sub Form_Load()          'Puts ups pictures to improve form appearance

picHargrave1.Picture = LoadPicture(App.Path & "\" & hargravepix(1))
picHargrave2.Picture = LoadPicture(App.Path & "\" & hargravepix(2))
picHargrave3.Picture = LoadPicture(App.Path & "\" & hargravepix(3))

End Sub
