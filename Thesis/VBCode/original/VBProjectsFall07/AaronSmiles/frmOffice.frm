VERSION 5.00
Begin VB.Form frmOffice 
   BackColor       =   &H8000000D&
   Caption         =   "Office Programs"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5490
   LinkTopic       =   "Form2"
   ScaleHeight     =   4860
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close Window"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   6
      Top             =   3720
      Width           =   3135
   End
   Begin VB.PictureBox picAccess 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   3720
      Picture         =   "frmOffice.frx":0000
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   5
      ToolTipText     =   "Access"
      Top             =   2040
      Width           =   1500
   End
   Begin VB.PictureBox picPowerPoint 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   1920
      Picture         =   "frmOffice.frx":1BC9
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   4
      ToolTipText     =   "PowerPoint"
      Top             =   2040
      Width           =   1500
   End
   Begin VB.PictureBox picOutlook 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   120
      Picture         =   "frmOffice.frx":32C3
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   3
      ToolTipText     =   "Outlook"
      Top             =   2040
      Width           =   1500
   End
   Begin VB.PictureBox picVisualBasic 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   3720
      Picture         =   "frmOffice.frx":4E21
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   2
      ToolTipText     =   "Visual Basic 6.0"
      Top             =   240
      Width           =   1500
   End
   Begin VB.PictureBox picExcel 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   1920
      Picture         =   "frmOffice.frx":6E38
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   1
      ToolTipText     =   "Excel"
      Top             =   240
      Width           =   1500
   End
   Begin VB.PictureBox picWord 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   120
      Picture         =   "frmOffice.frx":8462
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   0
      ToolTipText     =   "Word"
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "frmOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PowerPointPath, ExcelPath, AccessPath, OutlookPath, VBPath, WordPath As String


Private Sub cmdexit_Click()
frmOffice.Hide


'This form access useful Office programs directly. Without the variable, the shell command always asked permission to open the .exe file

End Sub

Private Sub picAccess_Click()
AccessPath = "C:\Program Files\Microsoft Office\Office12\MSACCESS.EXE"
Shell AccessPath
End Sub

Private Sub picExcel_Click()
ExcelPath = "C:\Program Files\Microsoft Office\Office12\EXCEL.EXE"
Shell ExcelPath
End Sub

Private Sub picOutlook_Click()
OutlookPath = "C:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE"
Shell OutlookPath
End Sub

Private Sub picPowerPoint_Click()
PowerPointPath = "C:\Program Files\Microsoft Office\Office12\POWERPNT.EXE"
Shell PowerPointPath
End Sub

Private Sub picVisualBasic_Click()
VBPath = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
Shell VBPath
End Sub

Private Sub picWord_Click()
WordPath = "C:\Program Files\Microsoft Office\Office12\WINWORD.EXE"
Shell WordPath
End Sub
