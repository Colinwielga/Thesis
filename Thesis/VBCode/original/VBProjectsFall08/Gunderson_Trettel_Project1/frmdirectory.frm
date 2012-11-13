VERSION 5.00
Begin VB.Form frmDirectory 
   BackColor       =   &H00008000&
   Caption         =   "Directory"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   3600
      Picture         =   "frmdirectory.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   2955
      TabIndex        =   13
      Top             =   8760
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   9000
      Picture         =   "frmdirectory.frx":77CA
      ScaleHeight     =   4395
      ScaleWidth      =   4755
      TabIndex        =   12
      Top             =   2040
      Width           =   4815
   End
   Begin VB.CommandButton cdmquit 
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   10320
      TabIndex        =   11
      Top             =   7680
      Width           =   2655
   End
   Begin VB.CommandButton cmdmiac 
      Caption         =   "GO!"
      Height          =   735
      Left            =   6240
      TabIndex        =   10
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdcalc 
      Caption         =   "GO!"
      Height          =   735
      Left            =   6240
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdteamr 
      Caption         =   "GO!"
      Height          =   735
      Left            =   6240
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdirgo 
      Caption         =   "GO!"
      Height          =   735
      Left            =   6240
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdccgo 
      Caption         =   "GO!"
      Height          =   735
      Left            =   6240
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "November 5, 2008"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   10200
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00008000&
      Caption         =   "Directory"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   10200
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00008000&
      Caption         =   "2008 MIAC Cross Country Project "
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   9960
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00008000&
      Caption         =   "By: Tyler Trettel and Josh Gunderson"
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   9960
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00008000&
      Caption         =   "MIAC Schools....................................................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   5895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      Caption         =   "Calculator............................................................"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   5895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      Caption         =   "What Is Cross Country......................................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Caption         =   "Team Results......................................................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   "Individual Results..............................................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   5895
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00008000&
      Caption         =   "Directory"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   975
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmdirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: MIAC CC Project
'Form Name: frmDirectory
'Authors: Josh Gunderson & Tyler Trettel
'Date: 5 November 2008
'Objective: The purpose of this form is for the user to navigate thru the program with ease.

Private Sub cdmquit_Click()

End

End Sub

Private Sub cmdcalc_Click()
frmdirectory.Hide
frmconversion.Show
End Sub

Private Sub cmdccgo_Click()
frmdirectory.Hide
frmExplaination.Show

End Sub

Private Sub cmdirgo_Click()


frmdirectory.Hide
frmIndivResults.Show

End Sub

Private Sub cmdmiac_Click()
frmdirectory.Hide
frmmiacschools.Show

End Sub

Private Sub cmdteamr_Click()
frmdirectory.Hide
frmTeamResults.Show

End Sub
