VERSION 5.00
Begin VB.Form frmFirstForm 
   BackColor       =   &H000000FF&
   Caption         =   "Rugby "
   ClientHeight    =   8535
   ClientLeft      =   3060
   ClientTop       =   2355
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   9045
   Begin VB.CommandButton cmdReadyButton 
      Caption         =   "Ready to Begin?"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   7320
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
End
Attribute VB_Name = "frmFirstForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdQuit_Click()
    End                 'Ends program
End Sub

Private Sub cmdReadyButton_Click()
    
    frmFirstForm.Hide       'removes form from screen
    frmSecondForm.Show      'pulls up the next form
End Sub


Private Sub Form_Load()
    Picture = LoadPicture("M:\CS130\miscellaneous\PROJECTS\RugbyPhoto.jpg") 'Projects picture onto the form
End Sub




                                                                        'Eric Bendzinski Project 1.vbp
                                                                        'frmFirstForm
                                                                        'Eric Bendzinski
                                                                        'Written 11/1/06 and 11/3/06
