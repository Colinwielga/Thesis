VERSION 5.00
Begin VB.Form frmOedipusTex 
   BackColor       =   &H80000007&
   Caption         =   "Oedipus Tex"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   Picture         =   "frmOedipusTex.frx":0000
   ScaleHeight     =   7485
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   3
      ToolTipText     =   "click to go back to previous slide."
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSynopsis 
      Caption         =   "Synopsis"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Click to see the plot of the opera."
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdCastList 
      Caption         =   "Cast List"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Click to go to form where it will show the cast list."
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lblOedipusTex 
      BackColor       =   &H80000007&
      Caption         =   "Oedipus Tex"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblDesigned 
      BackColor       =   &H80000007&
      Caption         =   "Designed By Amanda Weis"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmOedipusTex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'create button to show form for the Cast List
Private Sub cmdCastList_Click()
    frmOedipusTex.Hide
    frmOedipusCastList.Show
End Sub
    'create button to go back to previous form
Private Sub cmdGoBack_Click()
    frmOedipusTex.Hide
    frmOPERA.Show
End Sub
    'create button where a message box will appear giving synopsis of the opera
Private Sub cmdSynopsis_Click()
    MsgBox "Oedipus Tex is a one-act opera written by P. D. Q. Bach as a spoof on teh original Oedipus Rex.  In the original story, Oedipus finds the love of his life and murders her husband in order to marry her.  The town the couple settles down in is then hit with the plague.  To find out why the town becomes plagued with plague, Oedipus and his wife visits a fortune teller who tells them the reason why everyone in town is dying.  It turns out that Oedipus's wife is actually his mother!  Oedipus murdered his father and took his mother as his lover.  IN a fit of disgust, Oedipus's mother rushes home and kills herself.  When Oedipus finds out what his mother/lover has done, he pokes out both of his eyes and banishes himself to the wilderness.  Oedipus Tex follows the same guidelines, but puts a Texan spin on teh entire thing.", , "Oedipus Tex Synopsis"
End Sub


