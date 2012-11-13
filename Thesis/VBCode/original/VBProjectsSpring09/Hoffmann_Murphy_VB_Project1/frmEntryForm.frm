VERSION 5.00
Begin VB.Form frmEntryForm 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   Picture         =   "frmEntryForm.frx":0000
   ScaleHeight     =   7890
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEntry 
      BackColor       =   &H00000080&
      Caption         =   "Click to learn more!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      MaskColor       =   &H00000080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Big Sky Resort"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
   End
End
Attribute VB_Name = "frmEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Big Sky Resort
'frmEntryForm
'Ryan Hoffmann and Jamison Murphy
'Written on March 19, 2009
'This Form was created as a starting page for the user in order to get
'started, and see a beautiful picture of the surrounding environment


'We are simply moving from the first form and entering into the program
Private Sub cmdEntry_Click()
    frmHomeForm.Show
    frmEntryForm.Hide
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub

