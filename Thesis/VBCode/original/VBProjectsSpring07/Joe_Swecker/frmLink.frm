VERSION 5.00
Begin VB.Form frmLink 
   BackColor       =   &H0000FFFF&
   Caption         =   "Link To School website"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000C000&
      Caption         =   "Back to main page"
      Height          =   975
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Click on the link the link below"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblURL 
      BackColor       =   &H0000C000&
      Caption         =   "http://www.isd47.org/schools/saukrapids-ricehighschool/index.php"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   2760
      Width           =   3615
   End
End
Attribute VB_Name = "frmLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdBack_Click()
frmLink.Hide
frmStormBball.Show
End Sub

Private Sub lblURL_DragDrop(Source As Control, X As Single, Y As Single)
    If Source Is lblURL Then
        With lblURL
            .Font.Underline = False
            .ForeColor = vbBlack
            ' Call ShellExecute(0&, vbNullString, "Mailto:" & .Caption, vbNullString, vbNullString, vbNormalFocus)
            Call ShellExecute(0&, vbNullString, .Caption, vbNullString, vbNullString, vbNormalFocus)
        End With
    End If
End Sub

Private Sub lblURL_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbLeave Then
        With lblURL
            .Drag vbEndDrag
            .Font.Underline = False
            .ForeColor = vbBlack
        End With
    End If
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblURL
        .ForeColor = vbBlue
        .Font.Underline = True
        .Drag vbBeginDrag
    End With
End Sub 'http://www.samspublishing.com/library/content.asp?b=STY_VB6_24hours&seqNum=117&rl=1

