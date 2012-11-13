VERSION 5.00
Begin VB.Form frmPrograms 
   Caption         =   "Programs"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Hide Apps"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   7080
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdShowapps 
      Caption         =   "Quick Apps"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   6240
      Width           =   3855
   End
   Begin VB.PictureBox Picture6 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   3795
      TabIndex        =   12
      Top             =   6960
      Width           =   3855
   End
   Begin VB.PictureBox cmdIE 
      Height          =   615
      Left            =   3240
      Picture         =   "frmPrograms.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   11
      ToolTipText     =   "Internet Explorer"
      Top             =   6240
      Width           =   615
   End
   Begin VB.PictureBox cmdWindowsMedia 
      Height          =   615
      Left            =   2400
      Picture         =   "frmPrograms.frx":1552
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   10
      ToolTipText     =   "Opens Windows Media Player"
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox cmdSolitaire 
      Height          =   615
      Left            =   1560
      Picture         =   "frmPrograms.frx":15CB0
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   9
      ToolTipText     =   "Opens Solitaire"
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox cmdmessenger 
      Height          =   615
      Left            =   720
      Picture         =   "frmPrograms.frx":2F902
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   8
      ToolTipText     =   "Opens MSN Messenger"
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox cmdAdobe 
      Height          =   615
      Left            =   0
      Picture         =   "frmPrograms.frx":40F23
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      ToolTipText     =   "Opens Adobe Photoshiop"
      Top             =   6240
      Width           =   615
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      Height          =   915
      Left            =   0
      Picture         =   "frmPrograms.frx":47E98
      ScaleHeight     =   855
      ScaleWidth      =   3825
      TabIndex        =   6
      ToolTipText     =   "Click For CSB/SJU Homepage"
      Top             =   4680
      Width           =   3885
   End
   Begin VB.CommandButton cmdNotes 
      Caption         =   "Quick Notes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   5880
      Width           =   3855
   End
   Begin VB.TextBox txtNotes 
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmdFolders 
      Caption         =   "Folders"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   3855
   End
   Begin VB.CommandButton cmdMusic 
      Caption         =   "Music"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   3855
   End
   Begin VB.CommandButton cmdWWW 
      Caption         =   "Bookmarks"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CommandButton cmdOffice 
      Caption         =   "Office"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      MaskColor       =   &H00400000&
      Picture         =   "frmPrograms.frx":48B45
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmPrograms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the main window of the program. The majority of the commands here hide/show windows and open applications
'Information on below code can be found on frmInternet
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim MessShow As Boolean, UserName As String

Private Sub cmdAdobe_Click()
Shell ("C:\Program Files\Adobe\Adobe InDesign CS3\inDesign.exe")
End Sub

Private Sub cmdFolders_Click()
frmFolders.Show
frmMusic.Hide
frmOffice.Hide
frmInternet.Hide

End Sub

Private Sub cmdContacts_Click()
frmInternet.Hide
frmOffice.Hide
frmMusic.Hide
frmFolders.Hide

End Sub

Private Sub cmdGames_Click()
frmInternet.Hide
frmOffice.Hide
frmMusic.Hide
frmFolders.Hide

End Sub


Private Sub cmdIE_Click()
Shell ("C:\Program Files\Internet Explorer\iexplore.exe")
End Sub

Private Sub cmdmessenger_Click()
Shell ("C:\Program Files\Messenger\msmsgs.exe")
End Sub

Private Sub cmdMusic_Click()
frmMusic.Show
frmOffice.Hide
frmInternet.Hide
frmFolders.Hide

End Sub

Private Sub cmdNotes_Click() 'This shows and hides the note form, if it is hidden, clicking on the csb logo will take you to the homepage.
If txtNotes.Visible = False Then
txtNotes.Visible = True
pic.Visible = False
Else
txtNotes.Visible = False
pic.Visible = True
End If

End Sub

Private Sub cmdOffice_Click()
frmOffice.Show
frmInternet.Hide
frmMusic.Hide
frmFolders.Hide

End Sub

Private Sub cmdShowapps_Click()
If MessShow = False Then
    MsgBox "Hover over the buttons to find out which Applications they are associated with."
    MessShow = True
End If
cmdShowapps.Visible = False
Command1.Visible = True
End Sub

Private Sub cmdSolitaire_Click()
Shell ("C:\windows\system32\sol.exe")
End Sub

Private Sub cmdWindowsMedia_Click()
Shell ("C:\Program Files\Windows Media Player\wmplayer.exe")
End Sub

Private Sub cmdWWW_Click()
frmInternet.Show
frmOffice.Hide
frmMusic.Hide
frmFolders.Hide

End Sub
'Below are a set of quick links to applications which I find useful


Private Sub Picture3_Click()
Shell ("C:\Program Files\Adobe\Photoshop Elements 4.0\Photoshop Elements 4.0.exe")
End Sub

Private Sub Command1_Click()
cmdShowapps.Visible = True
Command1.Visible = False
End Sub

Private Sub Form_Load()
MessShow = False
UserName = InputBox("Please enter your name")
Label1.Caption = "Welcome " & UserName
End Sub



Private Sub pic_Click()
ShellExecute Me.hwnd, vbNullString, "http://www.csbsju.edu", vbNullString, "C:\", SW_SHOWNORMAL
End Sub


Private Sub Picture5_Click()
Shell ("C:\Program Files\Internet Explorer\iexplore.exe")
End Sub

