VERSION 5.00
Begin VB.Form frmWhatsRugby 
   BackColor       =   &H00404000&
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   3900
   ClientTop       =   3300
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6615
   Begin VB.CommandButton Command1 
      Caption         =   "Menu"
      Height          =   855
      Left            =   2160
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Video"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Best Rugby Tries"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Best rugby tackles"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "History of rugby"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Description of Positions"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click on these links to learn more about rugby."
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      Caption         =   "Rules of the Game"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
End
Attribute VB_Name = "frmWhatsRugby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. John's Rugby
'Sam Herrmann
'March 2009

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()

frmWhatsRugby.Hide
frmMenu.Show

End Sub

Private Sub Label2_Click()
    ShellExecute hWnd, "open", "http://rugby.union.rpi.edu/index_positions.html", vbNullString, vbNullString, conSwNormal

End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.FontBold = True
    Label2.FontUnderline = True
    Label2.ForeColor = vbBlue
    Me.MousePointer = 99
End Sub
'Private Sub Form_MouseMove2(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Label2.FontBold = False
    'Label2.FontUnderline = False
    'Label2.ForeColor = vbBlack
    'Me.MousePointer = 0
'End Sub
Private Sub Label3_Click()
    ShellExecute hWnd, "open", "http://www.rl1908.com/rugby-history.htm", vbNullString, vbNullString, conSwNormal
    
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.FontBold = True
    Label3.FontUnderline = True
    Label3.ForeColor = vbBlue
    Me.MousePointer = 99
End Sub

Private Sub Label4_Click()
    ShellExecute hWnd, "open", "http://www.youtube.com/watch?v=xvMFHXcd0yQ", vbNullString, vbNullString, conSwNormal

End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.FontBold = True
    Label4.FontUnderline = True
    Label4.ForeColor = vbBlue
    Me.MousePointer = 99
End Sub
Private Sub Label5_Click()
    ShellExecute hWnd, "open", "http://www.youtube.com/watch?v=QCLe2B2FY-o", vbNullString, vbNullString, conSwNormal

End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.FontBold = True
    Label5.FontUnderline = True
    Label5.ForeColor = vbBlue
    Me.MousePointer = 99
End Sub
Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLink.FontBold = True
    lblLink.FontUnderline = True
    lblLink.ForeColor = vbBlue
    Me.MousePointer = 99
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLink.FontBold = False
    lblLink.FontUnderline = False
    lblLink.ForeColor = vbBlack
    Me.MousePointer = 0
End Sub

Private Sub LblLink_Click()
    ShellExecute hWnd, "open", "http://www.ombac.org/ombac_rugby/rulesofrugby.htm", vbNullString, vbNullString, conSwNormal

End Sub
