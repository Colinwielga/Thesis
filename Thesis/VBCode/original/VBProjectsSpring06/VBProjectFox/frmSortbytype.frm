VERSION 5.00
Begin VB.Form frmSortbytype 
   BackColor       =   &H00000000&
   Caption         =   "Sort the trees by Species"
   ClientHeight    =   8475
   ClientLeft      =   1500
   ClientTop       =   1500
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11100
   Begin VB.CommandButton cmdWriteSort 
      Caption         =   "Click Here to write sorted data into a new file"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortbySpecies 
      Caption         =   "Click Here to Sort All Trees in File by Species"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   7
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdShowFile 
      Caption         =   "Click Here to Show File"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdSpecificSpecies 
      Caption         =   "Click Here to Find all instances of a particular species in your file"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoadFileSS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start by Clicking Here to Load File"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdRetfromSortType 
      BackColor       =   &H8000000B&
      Caption         =   "Return to First Slide"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8400
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortTypeEnd 
      Caption         =   "To End Program Click Here"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      TabIndex        =   1
      Top             =   6480
      Width           =   1575
   End
   Begin VB.PictureBox picSortResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   2520
      ScaleHeight     =   7275
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   840
      Width           =   5655
      Begin VB.VScrollBar VSDownSpecies 
         Height          =   7215
         Left            =   5280
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Image imgBackSbS 
      Height          =   7965
      Left            =   -480
      Picture         =   "frmSortbytype.frx":0000
      Top             =   480
      Width           =   12000
   End
   Begin VB.Label lblSbS 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort the Trees by Species"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmSortbytype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmSortbyType(frmSortbyType.frm)
'Author: Kelly Fox
'Date Written:3/22/2006
'This is a form allows the user to sort a file they input by species and then load their data into a new file
Option Explicit
Private Sub cmdLoadFileSS_Click()
    'Loads file into program
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

Private Sub cmdRetfromSortType_Click()
    frmSortbytype.Hide
    frmMinnesotaTrees.Show
End Sub

Private Sub cmdShowFile_Click()
'Displays file in orginal format on screen
    Dim Pos As Integer
    picSortResults.Cls
    picSortResults.Print "Position", "Common", , "Scientific"
    For Pos = 1 To Size
        picSortResults.Print Position(Pos), CommonName(Pos), SciName(Pos)
    Next Pos
End Sub

Private Sub cmdSortbySpecies_Click()
'sorts the scientific name of the species alphabetically along with parallel arrays
    Dim Pos As Integer
    Dim SciTemp As String
    Dim Pass As Integer
    Dim CommonTemp As String
    Dim PositionTemp As String
    picSortResults.Cls
    picSortResults.Print "Position", "Common", , "Scientific"
    'labels the information
    For Pass = 1 To (Size - 1)
        For Pos = 1 To (Size - Pass)
            If SciName(Pos) > SciName(Pos + 1) Then
                SciTemp = SciName(Pos)
                SciName(Pos) = SciName(Pos + 1)
                SciName(Pos + 1) = SciTemp
                'sort the scientific names alphabetically
                CommonTemp = CommonName(Pos)
                CommonName(Pos) = CommonName(Pos + 1)
                CommonName(Pos + 1) = CommonTemp
                'sorts the parallel array CommonName
                PositionTemp = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = PositionTemp
                'sorts the parallel array Position
            End If
        Next Pos
     Next Pass
     For Pos = 1 To Size
        picSortResults.Print Position(Pos), CommonName(Pos), SciName(Pos)
    Next Pos
End Sub

Private Sub cmdSortTypeEnd_Click()
    End
End Sub

Private Sub cmdSpecificSpecies_Click()
    'obtains search value from user and find  the number of species of that type
    Dim SV2 As String
    Dim Pos As Integer
    SV2 = InputBox("Type in the common name of species you wish to find", "Type in common name")
    picSortResults.Cls
    picSortResults.Print "Position", "Common", , "Scientific"
    For Pos = 1 To Size
        If InStr(CommonName(Pos), SV2) Then
            picSortResults.Print Position(Pos), CommonName(Pos), SciName(Pos)
        End If
    Next Pos
End Sub

Private Sub cmdWriteSort_Click()
    'Writes sorted data into new file
    Dim Pos As Integer
    Pos = 0
    Open App.Path & "\SortTrees.txt" For Output As #3
        For Pos = 1 To Size
            Write #3, Position(Pos), CommonName(Pos), SciName(Pos)
        Next Pos
    Close #3
End Sub
