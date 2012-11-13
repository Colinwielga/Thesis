VERSION 5.00
Begin VB.Form frmVote 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15525
   FillColor       =   &H000000C0&
   FillStyle       =   0  'Solid
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   15525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Start Form"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11880
      TabIndex        =   10
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdVote 
      Caption         =   "Vote"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Frame fmeOptions 
      BackColor       =   &H00000080&
      Height          =   2655
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   11535
      Begin VB.OptionButton optLiriano 
         BackColor       =   &H00000080&
         Caption         =   "Francisco Liriano"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   6
         Left            =   8160
         TabIndex        =   9
         Top             =   1200
         Width           =   2775
      End
      Begin VB.OptionButton optYoung 
         BackColor       =   &H00000080&
         Caption         =   "Delman Young"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   8160
         TabIndex        =   8
         Top             =   480
         Width           =   2655
      End
      Begin VB.OptionButton optCuddyer 
         BackColor       =   &H00000080&
         Caption         =   "Michael Cuddyer"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   3
         Left            =   4200
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton optMorneau 
         BackColor       =   &H00000080&
         Caption         =   "Justin Morneau"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   480
         TabIndex        =   6
         Top             =   1080
         Width           =   3015
      End
      Begin VB.OptionButton optSpan 
         BackColor       =   &H00000080&
         Caption         =   "Denard Span"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.OptionButton optOther 
         BackColor       =   &H00000080&
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   4800
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton optMauer 
         BackColor       =   &H00000080&
         Caption         =   "Joe Mauer"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00000080&
      Caption         =   "Click on one of the names below to cast your vote:"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   10935
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  'This form uses option buttons to allow the user to vote for their favorite player

Private Sub cmdReturn_Click() 'Return to the start form
    frmVote.Hide
    frmStart.Show
End Sub


Private Sub cmdVote_Click() 'After the user selects their favorite player and clicks "vote", display a message thanking them for voting

'Display message corresponding to the user's player choice
If optCuddyer(3).Value = True Then  'user selects Cuddyer
        MsgBox "Thank you for voting for Michael Cuddyer!"
    ElseIf optMauer(1).Value = True Then    'user selects Mauer
       MsgBox "Thank you for voting for Joe Mauer!"
    ElseIf optMorneau(2).Value = True Then  'user selects Morneau
        MsgBox "Thank you for voting for Justin Morneau!"
    ElseIf optSpan(4).Value = True Then 'user selects Span
        MsgBox "Thank you for voting for Denard Span!"
    ElseIf optYoung(5).Value = True Then    'user selects Young
        MsgBox "Thank you for voting for Delman Young!"
    ElseIf optLiriano(6).Value = True Then  'user selects Liriano
        MsgBox "Thank you for voting for Francisco Liriano!"
    ElseIf optOther(7).Value = True Then    'user selects the other option
        MsgBox "Thank you for voting!"
End If

End Sub

