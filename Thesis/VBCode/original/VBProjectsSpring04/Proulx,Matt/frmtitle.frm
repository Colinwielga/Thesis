VERSION 5.00
Begin VB.Form frmtitle 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Welcome to the Counting Crows Information Program"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictureDisplayBox 
      Height          =   4575
      Left            =   1080
      Picture         =   "frmtitle.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   8715
      TabIndex        =   5
      Top             =   5520
      Width           =   8775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   1680
      Picture         =   "frmtitle.frx":A691A
      ScaleHeight     =   3255
      ScaleWidth      =   7455
      TabIndex        =   4
      Top             =   120
      Width           =   7455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1335
      Left            =   8760
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdsale 
      Caption         =   "Counting Crows Store"
      Height          =   1335
      Left            =   5760
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdfans 
      Caption         =   "Fans"
      Height          =   1335
      Left            =   3000
      TabIndex        =   1
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmddics 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Discography"
      Height          =   1335
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Matt Proulx"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   10320
      Width           =   1695
   End
End
Attribute VB_Name = "frmtitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : CountingCrows (Matt Proulx's Counting Crows Program.vbp)
'Form Name : frmtitle (frmtitle.frm)
'Author: Matt Proulx
'Date Written: March 13, 2004
'Purpose of the Project: 'To introduce people to the band Counting Crows and what they are all about. It is also for
                         'current fans that just want get in touch with other fans of the band and purchase band items.
'Purpose of the Form:    'This form is simply the title form where the user can choose what part
                         'of the program they want to go to.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmddics_Click()
    frmtitle.Hide
    frmdisc.Show
End Sub
Private Sub cmdfans_Click()
    frmtitle.Hide
    frmfans.Show
End Sub
Private Sub cmdsale_Click()
    frmtitle.Hide
    frmstore.Show
End Sub
Private Sub cmdQuit_Click()
    End 'Quits the program
End Sub

