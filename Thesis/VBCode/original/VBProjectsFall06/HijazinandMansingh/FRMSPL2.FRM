VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11430
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMSPL~2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11430
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      ForeColor       =   &H00FFFF00&
      Height          =   11355
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   15105
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5400
         Top             =   6960
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8040
         TabIndex        =   5
         Top             =   6960
         Width           =   1275
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright: Meera and Ghaith"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   4
         Top             =   7440
         Width           =   2415
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Zone"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   48
            Charset         =   161
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1545
         Left            =   5760
         TabIndex        =   3
         Top             =   4800
         Width           =   3930
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Library"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6480
         TabIndex        =   2
         Top             =   4080
         Width           =   1665
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LicenseTo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   1
         Top             =   4200
         Width           =   975
      End
      Begin VB.Image imgLogo 
         Height          =   3465
         Left            =   720
         Picture         =   "FRMSPL~2.frx":000C
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub



Private Sub Timer1_Timer()
   frmMain.show
   
   Unload Me
End Sub
