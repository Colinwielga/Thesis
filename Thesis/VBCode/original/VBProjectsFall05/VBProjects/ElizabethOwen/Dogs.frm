VERSION 5.00
Begin VB.Form frmFirstscreen 
   BackColor       =   &H0000FF00&
   Caption         =   "Info on Dogs"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "MS Reference Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help me find the right breed"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      TabIndex        =   7
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdKnow 
      BackColor       =   &H0000FF00&
      Caption         =   "I already know the breed I want"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   3600
      TabIndex        =   6
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FF00&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   6240
      Width           =   2295
   End
   Begin VB.PictureBox picBull 
      Height          =   3495
      Left            =   3600
      Picture         =   "Dogs.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.PictureBox picLab 
      Height          =   2415
      Left            =   6480
      Picture         =   "Dogs.frx":4DA3
      ScaleHeight     =   2355
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
   End
   Begin VB.PictureBox picdal 
      Height          =   3255
      Left            =   240
      Picture         =   "Dogs.frx":D4E7
      ScaleHeight     =   3195
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H0000FF00&
      Caption         =   "When looking for the perfect dog there are many things that one should take into account "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   720
      TabIndex        =   4
      Top             =   4680
      Width           =   8055
   End
   Begin VB.Label lblIntro 
      BackColor       =   &H0000FF00&
      Caption         =   "How to Choose the Best Dog for You"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "frmFirstscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Dogs(VB-project.vbp)
'Form Name: frmFirstscreen (FrmDogs.frm)
'Author: Libby Owen
'Date: Wednesday October 19
'Purpose of the project: to have the user interact with the program to decide
                        ' what kind of dog would be best for them.  The program
                        'will educate the user on some common types of dogs and
                        ' what to look for when picking one to take home
'Purpose of the form: this is the first page seen, the user will have to use this
                        ' page before they can continue on with the program. The
                        'user will have to click on a button that best fits what
                        ' they want to do.
        
Private Sub cmdExit_Click()
End 'allows a user to end the program

End Sub


Private Sub cmdHelp_Click()
    frmFind.Show 'goes to the page to help user pick a breed
    frmFirstscreen.Hide
End Sub

Private Sub cmdKnow_Click()
    frmBreeds.Show 'goes to the form where the user can look at info on the breed they
    frmFirstscreen.Hide ' chose.
    
                
End Sub


