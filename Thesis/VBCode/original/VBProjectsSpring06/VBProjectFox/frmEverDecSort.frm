VERSION 5.00
Begin VB.Form frmEverDecSort 
   BackColor       =   &H00000040&
   Caption         =   "Sort Into Deciduous and Evergreen Trees"
   ClientHeight    =   8355
   ClientLeft      =   1500
   ClientTop       =   2130
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   Picture         =   "frmEverDecSort.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   11730
   Begin VB.PictureBox picUnknown 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      ScaleHeight     =   1755
      ScaleWidth      =   7275
      TabIndex        =   9
      Top             =   6480
      Width           =   7335
      Begin VB.VScrollBar VScroll3 
         Height          =   1815
         Left            =   6960
         TabIndex        =   13
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picEvergreen 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   5655
      Left            =   4320
      ScaleHeight     =   5595
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   360
      Width           =   4215
      Begin VB.VScrollBar VScroll2 
         Height          =   5655
         Left            =   3960
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picDecid 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5595
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      Begin VB.VScrollBar VScroll1 
         Height          =   5655
         Left            =   3720
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000040&
      Caption         =   "Unknown Classification"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label lblEndProgram 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Click Below to End Program"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   8880
      TabIndex        =   8
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label lblReturnDE 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Click Below to Return to First Slide"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   8760
      TabIndex        =   7
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000040&
      Caption         =   "           Deciduous Trees"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image imgReturnfromED 
      Height          =   975
      Left            =   9120
      Picture         =   "frmEverDecSort.frx":50B2
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1530
   End
   Begin VB.Image imgEndEverDec 
      Height          =   1275
      Left            =   9120
      Picture         =   "frmEverDecSort.frx":B0EC
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1620
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000040&
      Caption         =   "           "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label lblLoadDE 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Click below to load your file "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image imgLoadDE 
      Height          =   960
      Left            =   9120
      Picture         =   "frmEverDecSort.frx":12772
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1425
   End
   Begin VB.Image imgDecidTrees 
      Height          =   975
      Left            =   9000
      Picture         =   "frmEverDecSort.frx":1ABB4
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1680
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   5880
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblFindDecid 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Click below to Sort Trees"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   8880
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000040&
      Caption         =   "Evergreen Trees"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmEverDecSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Identifying and Organizing sets of Trees from Minnesota
'frmEverDecSortfrmEverDecSort.frm)
'Author: Kelly Fox
'Date Written:3/22/2006
'Allows user to sort the trees in a file they upload into basic deciduous and everygreen categories
Private Sub Form_Load()
    frmEverDecSort.Picture = LoadPicture("")
End Sub

Private Sub imgDecidTrees_Click()
    'Splits trees in deciduous and evergreen and prints them in correct column
    Dim Pos As Double
    picDecid.Cls
    picEvergreen.Cls
    picUnknown.Cls
    picEvergreen.Print "Position", "Common", , "Scientific"
    picDecid.Print "Position", "Common", , "Scientific"
    picUnknown.Print "Positon", "Common", , "Scientific"
    For Pos = 1 To Size
        If InStr(SciName(Pos), "Quercus") <> 0 Or InStr(SciName(Pos), "Acer") <> 0 Or InStr(SciName(Pos), "Fraxinus") <> 0 Or InStr(SciName(Pos), "Ulmus") <> 0 Or InStr(SciName(Pos), "Ostrya") <> 0 Or InStr(SciName(Pos), "Betula") <> 0 Then
            picDecid.Print Position(Pos), CommonName(Pos), SciName(Pos)
        ElseIf InStr(SciName(Pos), "Pinus") <> 0 Or InStr(SciName(Pos), "Picea") <> 0 Then
            picEvergreen.Print Position(Pos), CommonName(Pos), SciName(Pos)
        Else
            picUnknown.Print Position(Pos), CommonName(Pos), SciName(Pos)
        End If
    Next Pos
End Sub

Private Sub imgEndEverDec_Click()
    End
End Sub

Private Sub imgLoadDE_Click()
    'Loads file from User
    Dim Pos As Double
    Open App.Path & "/TreesList.txt" For Input As #1
    Size = 0
    Pos = 0
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Position(Pos), CommonName(Pos), SciName(Pos)
    Loop
    Close #1
    Size = Pos
End Sub

Private Sub imgReturnfromED_Click()
    frmEverDecSort.Hide
    frmMinnesotaTrees.Show
End Sub


