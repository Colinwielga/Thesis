VERSION 5.00
Begin VB.Form frmOPERA 
   BackColor       =   &H00000000&
   Caption         =   "OPERA 2006"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   Picture         =   "frmOPERA.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      TabIndex        =   2
      ToolTipText     =   "Click to Exit"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOedipusTex 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Oedipus Tex"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      ToolTipText     =   "Click here to learn more about Oedipus Tex."
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdTrialByJury 
      Caption         =   "Trial By Jury"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H0000C000&
      TabIndex        =   0
      ToolTipText     =   "Click here to learn more about Trial By Jury."
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblOpera2006 
      BackColor       =   &H00000000&
      Caption         =   "Opera 2006"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   720
      TabIndex        =   4
      Top             =   6240
      Width           =   5295
   End
   Begin VB.Label lblDesign 
      BackColor       =   &H00000000&
      Caption         =   "Designed by Amanda Weis"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   7680
      TabIndex        =   3
      Top             =   6960
      Width           =   1335
   End
End
Attribute VB_Name = "frmOPERA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'create a button to end entire program
Private Sub cmdExit_Click()
    End
End Sub
    'create a button to take user to desired form, either Oedipus Tex, or Trial By Jury
    'create button to show Oedipus Tex form
Private Sub cmdOedipusTex_Click()
    frmOedipusTex.Show
    frmOPERA.Hide
End Sub
    'create button to show Trial By Jury Form
Private Sub cmdTrialByJury_Click()
    frmTrialByJury.Show
    frmOPERA.Hide
End Sub

