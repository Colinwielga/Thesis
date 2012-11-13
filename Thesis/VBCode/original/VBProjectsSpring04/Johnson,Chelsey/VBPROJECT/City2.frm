VERSION 5.00
Begin VB.Form City2 
   BackColor       =   &H0080FFFF&
   Caption         =   "City 2"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      TabIndex        =   9
      Top             =   10320
      Width           =   1095
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13680
      TabIndex        =   8
      Top             =   10320
      Width           =   1935
   End
   Begin VB.CommandButton cmdchoice 
      Caption         =   "Click here to find out if we have the same choice for which site we would rather see."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13560
      TabIndex        =   7
      Top             =   8880
      Width           =   3375
   End
   Begin VB.TextBox txtchoice 
      Height          =   615
      Left            =   10080
      TabIndex        =   6
      Top             =   8880
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   3240
      Picture         =   "City2.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   5880
      Width           =   5415
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   3720
      Picture         =   "City2.frx":68FD
      ScaleHeight     =   2955
      ScaleWidth      =   13155
      TabIndex        =   1
      Top             =   2400
      Width           =   13215
   End
   Begin VB.Label Label5 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   11760
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Which site would you perfer to see?  ""Guildhall "" or ""Museum of London"", Please enter one of these names."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Top             =   8880
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "This popular museum has an extensive display of london life, archeology and artefacts. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   4
      Top             =   6600
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "The Guildhall and The Museum of London are very important stops in the City 2 District of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   12015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"City2.frx":CBD3
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   0
      Top             =   1440
      Width           =   5895
   End
End
Attribute VB_Name = "City2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: City 2 (City2.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: The purpose of this form is to familiarize the user with The Guildhall and The Museum of London.
                    'They are able to vote for which one they would rather see and then learn if it was the same'
                    'choice as mine
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdchoice_Click()
Dim Choice As String
Choice = txtchoice.Text 'Getting the variable from the user, so it can later be compared
If Choice = "Guildhall" Then  'if true it will print out that we have different choices
    MsgBox "We do not have the same first choice.  I would rather visit the Museum of London first.", , "Guildhall"
End If
If Choice = "Museum of London" Then 'if true it will print out that we have the same choices
    MsgBox "Yippe!  We have the same first choice.  I would also choose the Museum of London first.", , "Museum of London"
End If
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'Returns the user back to the Map of London page so they are able to choose another district to look at
City2.Hide
MapLondon.Show
End Sub
