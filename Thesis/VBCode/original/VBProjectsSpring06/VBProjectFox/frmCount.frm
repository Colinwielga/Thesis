VERSION 5.00
Begin VB.Form frmCount 
   BackColor       =   &H00400000&
   Caption         =   "Count the Number of Trees"
   ClientHeight    =   9570
   ClientLeft      =   1290
   ClientTop       =   1080
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Nueva Std"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   11160
   Begin VB.PictureBox picCountTrees 
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8355
      ScaleWidth      =   4515
      TabIndex        =   7
      Top             =   120
      Width           =   4575
      Begin VB.VScrollBar VScroll1 
         Height          =   8415
         Left            =   4320
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   120
      Top             =   0
      Width           =   2895
   End
   Begin VB.Image imgAll 
      Height          =   1005
      Left            =   6960
      Picture         =   "frmCount.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1800
   End
   Begin VB.Label imgTitleC 
      BackColor       =   &H00400000&
      Caption         =   "Start by Loading the Program: By Clicking Below"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   120
      Width           =   6495
   End
   Begin VB.Image imgEndCount 
      Height          =   1350
      Left            =   8400
      Picture         =   "frmCount.frx":C8E2
      Top             =   7320
      Width           =   2025
   End
   Begin VB.Label lblEndC 
      BackColor       =   &H00400000&
      Caption         =   "Click Below to end program"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   8400
      TabIndex        =   5
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Image imgCountTotal 
      Height          =   1620
      Left            =   8760
      Picture         =   "frmCount.frx":15894
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgLoadforCount 
      Height          =   1275
      Left            =   7080
      Picture         =   "frmCount.frx":1D0F6
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1395
   End
   Begin VB.Image imgCountSpecific 
      Height          =   1575
      Left            =   5640
      Picture         =   "frmCount.frx":25CE0
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00400000&
      Caption         =   "If you want the total number of all species click below"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   7920
      TabIndex        =   4
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label lblAllMinn 
      BackColor       =   &H00400000&
      Caption         =   "If you want the number of all typical l Minnesota species click below"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   6720
      TabIndex        =   3
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label lblCertain 
      BackColor       =   &H00400000&
      Caption         =   "If you want the numberof a certain tree species click below"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   5160
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label lblTask 
      BackColor       =   &H00400000&
      Caption         =   "Choose the task you want to preform"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Label lblReturnfromC 
      BackColor       =   &H00400000&
      Caption         =   "Click Trees Below to Return to First Slide"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Image imgReturnfromCount 
      Height          =   1440
      Left            =   5400
      Picture         =   "frmCount.frx":2F24A
      Top             =   7320
      Width           =   1920
   End
End
Attribute VB_Name = "frmCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmCount(frmCount.frm)
'Author: Kelly Fox
'Date Written:3/21/2006
'This form allows the user to count the number of trees of particular species in a file uploaded by the user as well as display the number of trees of common Minnesota species in the file
Option Explicit

Private Sub imgLoadforCount_Click()
    'Loads file from user
    Dim Pos As Double
    Size = 0
    Pos = 0
    Open App.Path & "/TreesList.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Position(Pos), CommonName(Pos), SciName(Pos)
    Loop
    Close #1
    Size = Pos
End Sub
Private Sub imgAll_Click()
    'Displays the amounts of common trees in Minnesota in the file
    Dim Pos As Double
    Dim Oaks As Integer
    Dim RedOaks As Integer
    Dim WhiteOaks As Integer
    Dim ScarletOaks As Integer
    Dim Pines As Integer
    Dim WhitePines As Integer
    Dim Elms As Integer
    Dim Maples As Integer
    Dim Birches As Integer
    Dim PaperBirches As Integer
    Dim Spruce As Integer
    Dim NorwaySpruce As Integer
    Dim SugarMaples As Integer
    Dim NorwayPine As Integer
    SugarMaples = 0
    Oaks = 0
    RedOaks = 0
    WhiteOaks = 0
    ScarletOaks = 0
    Pines = 0
    WhitePines = 0
    Spruce = 0
    NorwaySpruce = 0
    NorwayPine = 0
    PaperBirches = 0
    picCountTrees.Cls
    picCountTrees.Print "Common", "Genus/Species", "# of Trees"
    'Labels the picture box so information is more understandable
    For Pos = 1 To Size
         If InStr(SciName(Pos), "Quercus") <> 0 Then
            Oaks = Oaks + 1
            'finds all the oaks in the file and adds them to total number of oaks to be printed
            If InStr(SciName(Pos), "borealis") Then
                RedOaks = RedOaks + 1
            ElseIf InStr(SciName(Pos), "alba") Then
                WhiteOaks = WhiteOaks + 1
            End If
         End If
    Next Pos
    For Pos = 1 To Size
         If InStr(SciName(Pos), "Acer") <> 0 Then
            Maples = Maples + 1
            If InStr(SciName(Pos), "saccharum") Then
                SugarMaples = SugarMaples + 1
            End If
         End If
    Next Pos
    For Pos = 1 To Size
         If InStr(SciName(Pos), "Pinus") <> 0 Then
            Pines = Pines + 1
            If InStr(SciName(Pos), "strobus") Then
                WhitePines = WhitePines + 1
            ElseIf InStr(SciName(Pos), "resinosa") Then
                NorwayPine = NorwayPine + 1
            End If
         End If
    Next Pos
    For Pos = 1 To Size
        If InStr(SciName(Pos), "Ulmus") <> 0 Then
            Elms = Elms + 1
        End If
    Next Pos
    For Pos = 1 To Size
        If InStr(SciName(Pos), "Betula") <> 0 Then
            Birches = Birches + 1
            If InStr(SciName(Pos), "papyrifera") Then
                PaperBirches = PaperBirches + 1
            End If
        End If
    Next Pos
    For Pos = 1 To Size
        If InStr(SciName(Pos), "Picea") <> 0 Then
            Spruce = Spruce + 1
            If InStr(SciName(Pos), "resinosa") Then
                NorwaySpruce = NorwaySpruce + 1
            End If
        End If
    Next Pos
    picCountTrees.Print "***********"; Tab(2); "Oaks-", "Quercus", , Oaks; Tab(2); "----------"; Tab(2); "Red Oaks-"; "Quercus borealis", RedOaks; Tab(2); "----------"; Tab(2); "White Oaks-"; "Quercus strobus", WhiteOaks
    picCountTrees.Print "***********"; Tab(2); "Spruces-", "Picea", , Spruce; Tab(2); "---------"; Tab(2); "Norway Spruce-"; "Picea resinosa", NorwaySpruce
    picCountTrees.Print "***********"; Tab(2); "Pines-", "Pinus", , Pines; Tab(2); "----------"; Tab(2); "White Pines-"; "    Pinus strobus", WhitePines; Tab(2); "--------"; Tab(2); "Norway Pines-"; "Pinus resinosa", NorwayPine
    picCountTrees.Print "***********"; Tab(2); "Elms-", "Ulmus", , Elms
    picCountTrees.Print "***********"; Tab(2); "Birches-", "Betula", , Birches; Tab(2); "-----------"; Tab(2); "Paper Birches-"; "Betula papyifera", PaperBirches
    picCountTrees.Print "***********"; Tab(2); "Maples", "Acer", , Maples; Tab(2); "-----------"; Tab(2); "Sugar Maples-"; "Acer saccharum", SugarMaples
End Sub


Private Sub imgCountSpecific_Click()
    'Seraches the file for SearchValue provided by user and counts the number of times it appears
    Dim SearchValue As String
    SearchValue = InputBox("Enter the common name of the tree you which to find (Remember to capitalize first letter of each word)", "Enter name of tree")
    Dim Pos As Integer
    Dim Counter As Integer
    Counter = 0
    picCountTrees.Cls
    For Pos = 1 To Size
        If InStr(CommonName(Pos), SearchValue) <> 0 Then
           Counter = Counter + 1
        End If
    Next Pos
    picCountTrees.Print "There are " & Counter & " " & SearchValue & " trees in your file"
End Sub

Private Sub imgEndCount_Click()
    End
End Sub
Private Sub imgCountTotal_Click()
    picCountTrees.Cls
    picCountTrees.Print " There are " & Size & " trees in your file"
End Sub

Private Sub imgReturnfromCount_Click()
    frmCount.Hide
    frmMinnesotaTrees.Show
End Sub

