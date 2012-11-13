VERSION 5.00
Begin VB.Form frmfacts 
   BackColor       =   &H00004080&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddrink 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Have a drink!"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CommandButton cmdbath 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Take a leak..."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   2655
   End
End
Attribute VB_Name = "frmfacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
    'Project name:  Tour De St. Joe
    'Form:  frmfacts, "Facts"
    'Author:  Josh
    'Date:  3/26/08
    'Objective: To get out of this form and back to the original page


Private Sub cmdbath_Click()

    MsgBox "A fight broke out...  EVERYBODY OUT!!!"
    frmfacts.Hide
    frmjoetown.Show
    
End Sub


Private Sub cmddrink_Click()

    MsgBox "A fight broke out...  EVERYBODY OUT!!!"
    frmfacts.Hide
    frmjoetown.Show
    
End Sub
