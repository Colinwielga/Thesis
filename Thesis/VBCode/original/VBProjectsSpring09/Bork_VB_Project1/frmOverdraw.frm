VERSION 5.00
Begin VB.Form frmOverdraw 
   Caption         =   "Wilderness Outfitters"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   945
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   Picture         =   "frmOverdraw.frx":0000
   ScaleHeight     =   6705
   ScaleWidth      =   4170
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous Page"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   6000
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   4335
      Left            =   360
      ScaleHeight     =   4275
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Items Overdrawn:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "I'm sorry!  The are not enough of the following items to meet your requests.  Check Inventory for availability."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmOverdraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wilderness Outfitters Partial Outfitting
'Justin Bork
'March, 2009
'
'frmOverdraw
'The purpose of this form is to inform the user of which items have been
'overdrawn.  If an item is unavailable at the user's requested quantity it is
'printed and shown in this form.

Private Sub cmdPrevious_Click()
    frmOverdraw.Hide
    frmStartup.Show
End Sub
