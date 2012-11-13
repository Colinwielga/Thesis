VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H8000000D&
   Caption         =   "Jeopardy Information"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Jeopardy!"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   4
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Type in the address above and click to go now!"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2160
      TabIndex        =   1
      Text            =   "http://"
      Top             =   2760
      Width           =   5655
   End
   Begin VB.Label lblGoTo 
      BackColor       =   &H80000013&
      Caption         =   $"frmInfo.frx":0000
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   7815
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: CSJ/SJU Jeopardy
'Form Name: frmInfo
'Authors: Emma Jaynes, Linsday Havlik, Brooke Beyer
'Date Written: 11/04/08
'Objective: This form allows the user to enter in a web address, specifically Jeopardy.com,
'   and take them to that site in a web browser after they press the specified button.
'Other Comments: The enter may enter in any web address, but we hope they will chose Jeopardy.com!
'   Also, there is a "Quit" button and "Back to Main Menu" button.


Option Explicit
'necessary code for loading a web browser

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim SW_SHOW As Boolean, SW_NORMAL As Boolean

Private Sub cmdMainMenu_Click()
'takes contestand back to main menu

frmInfo.Hide
frmMainMenu.Show

End Sub

Private Sub cmdQuit_Click()
End

End Sub

Private Sub Command1_Click()
'takes the contestant to the web page once clicked

Dim URL As String

URL = txtURL.Text

ShellExecute Me.hWnd, "open", URL, "", "", SW_SHOW Or SW_NORMAL

End Sub


