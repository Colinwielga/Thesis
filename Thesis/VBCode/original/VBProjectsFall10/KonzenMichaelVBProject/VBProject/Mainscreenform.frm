VERSION 5.00
Begin VB.Form frmmainscreen 
   Caption         =   "Main Screen"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   Picture         =   "Mainscreen form.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      TabIndex        =   6
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdhalloffame 
      Caption         =   "Go to UFC Hall of Fame"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdweightclass 
      Caption         =   "Look Up Weight Classes"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdschedule 
      Caption         =   "Schedule"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      TabIndex        =   3
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdsponsor 
      Caption         =   "Search 2010 UFC Sponsors"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdreferee 
      Caption         =   "Meet the Referees"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   1
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdhistory 
      Caption         =   "Go to History of UFC"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmmainscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'all of these buttons on the main screen will bring you to a different frame that you can explore
'they will also hide the main screen when you go to a different frame
Private Sub cmdhalloffame_Click()
frmhalloffame.Show
frmmainscreen.Hide
End Sub

Private Sub cmdhistory_Click()
frmhistory.Show
frmmainscreen.Hide
End Sub

'this button ends the program
Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreferee_Click()
frmreferee.Show
frmmainscreen.Hide
End Sub

Private Sub cmdschedule_Click()
frmschedule.Show
frmmainscreen.Hide
End Sub

Private Sub cmdsponsor_Click()
frmsponsor.Show
frmmainscreen.Hide
End Sub

Private Sub cmdweightclass_Click()
frmweightclass.Show
frmmainscreen.Hide
End Sub

