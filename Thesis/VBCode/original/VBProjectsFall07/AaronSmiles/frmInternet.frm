VERSION 5.00
Begin VB.Form frmInternet 
   BackColor       =   &H8000000D&
   Caption         =   "Internet"
   ClientHeight    =   7830
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4380
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form3"
   ScaleHeight     =   7830
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLink3 
      Cancel          =   -1  'True
      Caption         =   "MSN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   2505
      TabIndex        =   16
      Top             =   600
      Width           =   2565
      Begin VB.CommandButton cmdLink1 
         Caption         =   "Google"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   6600
      Width           =   2535
   End
   Begin VB.TextBox txtText4 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "bbc.com"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK4 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset4 
      Caption         =   "Change Address"
      Height          =   855
      Left            =   2760
      TabIndex        =   12
      Top             =   5520
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   2475
      TabIndex        =   11
      Top             =   5520
      Width           =   2535
      Begin VB.CommandButton cmdLink4 
         Caption         =   "BBC"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK3 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset3 
      Caption         =   "Change  Address"
      Height          =   855
      Left            =   2760
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK2 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset2 
      Caption         =   "Change Address"
      Height          =   855
      Left            =   2760
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtText1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "google.com"
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdReset1 
      Caption         =   "Change Address"
      Height          =   855
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtText3 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "msn.com"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtText2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "apple.com"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   3960
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
      Begin VB.CommandButton cmdLink2 
         Caption         =   "Apple"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Menu File 
      Caption         =   ""
   End
End
Attribute VB_Name = "frmInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form provides quick links to useful websites. the links can be changed at any time.
Option Explicit
Dim URL(1 To 4) As String, picURL(1 To 4) As String
'the below is not my code(only under Option explicit and some of the ShellExecute Command code). While I came to understand some of it in the course of making this program,
'the actual code itself was taken from this website "http://www.codeguru.com/forum/archive/index.php/f-4-p-196.html"
'I modified it for my own purposes.

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1



Private Sub cmdexit_Click() 'exit program
frmInternet.Hide
End Sub



Private Sub cmdOK1_Click()
    txtText1.Visible = False 'hides the textbox until the user wishes to change the URL again
        cmdReset1.Visible = True    'shows the reset button again for the user
            cmdOK1.Visible = False  'hides the Ok box until user hit reset
End Sub

Private Sub cmdOK2_Click()
    txtText2.Visible = False
        cmdReset2.Visible = True
            cmdOK2.Visible = False
End Sub

Private Sub cmdOK3_Click()
txtText3.Visible = False
cmdReset3.Visible = True
cmdOK3.Visible = False
End Sub

Private Sub cmdOK4_Click()
txtText4.Visible = False
cmdReset4.Visible = True
cmdOK4.Visible = False
End Sub

Private Sub cmdReset2_Click()
txtText2.Visible = True 'makes the textbx visible for editing
cmdReset2.Visible = False   'hides the reset button
cmdOK2.Visible = True   'Ok button becomes visible
MsgBox "Enter a URL and press OK. http:\\www. has been added for you already."  'informs the user that http:// is not necessary to add into
End Sub

Private Sub cmdReset3_Click()
txtText3.Visible = True
cmdReset3.Visible = False
cmdOK3.Visible = True
MsgBox "Enter a URL and press OK. http:\\www. has been added for you already."
End Sub

Private Sub cmdReset4_Click()
txtText4.Visible = True
cmdReset4.Visible = False
cmdOK4.Visible = True
MsgBox "Enter a URL and press OK. http:\\www. has been added for you already."
End Sub


Private Sub cmdLink1_Click()
picURL(1) = txtText1.Text   '
If txtText1.Text <> "google.com" Then
cmdLink1.Caption = txtText1.Text
End If
URL(1) = picURL(1)
ShellExecute Me.hwnd, vbNullString, "http://www." & URL(1), vbNullString, "C:\", SW_SHOWNORMAL

End Sub


Private Sub cmdReset1_Click()
txtText1.Visible = True
cmdReset1.Visible = False
cmdOK1.Visible = True
MsgBox "Enter a URL and press OK. http:\\www. has been added for you already."
End Sub

Private Sub cmdLink2_Click()
picURL(2) = txtText2.Text   'variable is equal to the text in the text boxes which is set to apple
If txtText2.Text <> "apple.com" Then    'checks if the default caption is still valid
cmdLink2.Caption = txtText2.Text    'if not then the caption is changed to the desired URL
End If
URL(2) = picURL(2)  'variable is changed so that it can be entered into code below properly
ShellExecute Me.hwnd, vbNullString, "http://www." & URL(2), vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub cmdLink3_Click()
picURL(3) = txtText3.Text
If txtText3.Text <> "msn.com" Then
cmdLink3.Caption = txtText3.Text
End If
URL(3) = picURL(3)
ShellExecute Me.hwnd, vbNullString, "http://www." & URL(3), vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub cmdLink4_Click()
picURL(4) = txtText4.Text
If cmdLink4.Caption <> "BBC" Then
cmdLink4.Caption = txtText4.Text
End If
URL(4) = picURL(4)
ShellExecute Me.hwnd, vbNullString, "http://www." & URL(4), vbNullString, "C:\", SW_SHOWNORMAL
End Sub

