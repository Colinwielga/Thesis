VERSION 5.00
Begin VB.Form Entree 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form3"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form3"
   ScaleHeight     =   4725
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitch2 
      BackColor       =   &H00C000C0&
      Caption         =   "Back to Main Form"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton cmdAsian2 
      BackColor       =   &H00C000C0&
      Caption         =   "Entree 2: Asian"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdItalian2 
      BackColor       =   &H00C000C0&
      Caption         =   "Entree 1: Italian"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblEntree 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Entree"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "Entree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAsian2_Click()
AsianEntree.Show    'this shows the asian entree form
Entree.Hide         'this hides the entree form
End Sub

Private Sub cmdItalian2_Click()
ItalianEntree.Show  'this shows the italian entree form
Entree.Hide         'this hides the entree form
End Sub

Private Sub cmdSwitch2_Click()
Main.Show       'this shows the main form
Entree.Hide     'this hides the entree form
End Sub
