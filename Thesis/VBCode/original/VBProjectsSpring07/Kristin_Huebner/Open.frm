VERSION 5.00
Begin VB.Form frmArtHistoryOpen 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Art of India"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8175
   DrawWidth       =   5
   FillColor       =   &H00000080&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MouseIcon       =   "Open.frx":0000
   ScaleHeight     =   3675
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00AC796C&
      Caption         =   "Exit Program"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetStarted 
      Appearance      =   0  'Flat
      BackColor       =   &H00AC796C&
      Caption         =   "Get Started"
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      MousePointer    =   4  'Icon
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Indian Art History Through the          Kushan Period: Survey and                     Review Program"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   6135
   End
End
Attribute VB_Name = "frmArtHistoryOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' Introduction page
Private Sub cmdGetStarted_Click()
    frmArtHistoryOpen.Visible = False
    frmUsr_Info.Visible = True
    
End Sub

Private Sub Command1_Click()
    End
End Sub

