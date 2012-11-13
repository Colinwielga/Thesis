VERSION 5.00
Begin VB.Form frmsignon 
   Appearance      =   0  'Flat
   BackColor       =   &H00F3F3F3&
   BorderStyle     =   0  'None
   Caption         =   "MiM: welcome"
   ClientHeight    =   3435
   ClientLeft      =   5955
   ClientTop       =   5175
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2550
      Left            =   560
      Picture         =   "signonscreenguest.frx":0000
      ScaleHeight     =   2550
      ScaleWidth      =   6150
      TabIndex        =   7
      Top             =   450
      Width           =   6150
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEB580&
         Caption         =   ":enter password"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2320
         TabIndex        =   12
         Top             =   2290
         Width           =   2195
      End
      Begin VB.Label lblusername 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEB580&
         Caption         =   ":enter username"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   2290
         Width           =   2205
      End
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F3F3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Outlook"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00288A6A&
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3000
      Width           =   2190
   End
   Begin VB.PictureBox picsep 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   5
      Top             =   3000
      Width           =   135
   End
   Begin VB.TextBox txtuser 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F3F3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00288A6A&
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   2190
   End
   Begin VB.Label lblclose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   6720
      TabIndex        =   13
      Top             =   0
      Width           =   540
   End
   Begin VB.Label lbli 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BA3A30&
      Height          =   735
      Index           =   1
      Left            =   165
      TabIndex        =   10
      Top             =   -280
      Width           =   495
   End
   Begin VB.Label lblm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00288A6A&
      Height          =   735
      Index           =   3
      Left            =   -55
      TabIndex        =   9
      Top             =   170
      Width           =   855
   End
   Begin VB.Label lblm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00288A6A&
      Height          =   735
      Index           =   2
      Left            =   400
      TabIndex        =   8
      Top             =   -280
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "sign-on"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DA6903&
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00288A6A&
      Height          =   735
      Index           =   1
      Left            =   6560
      TabIndex        =   4
      Top             =   2270
      Width           =   855
   End
   Begin VB.Label lblm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00288A6A&
      Height          =   735
      Index           =   0
      Left            =   5990
      TabIndex        =   3
      Top             =   2720
      Width           =   855
   End
   Begin VB.Label lbli 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BA3A30&
      Height          =   735
      Index           =   0
      Left            =   6870
      TabIndex        =   0
      Top             =   2722
      Width           =   495
   End
End
Attribute VB_Name = "frmsignon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim change As Integer, change2 As Integer

Private Sub Label1_Click()
    Dim user As String
    Dim password As String
    Dim users(1 To 2) As String, passwords(1 To 2) As String
    Dim count As Integer, correct As Integer, count2
    user = txtuser.Text
    password = txtpassword.Text
    Open App.Path & "\users.txt" For Input As #1
        For count = 1 To 2
            Input #1, users(count), passwords(count)
        Next count
    Close #1
    count = 0
    For count = 1 To 2
        If user = users(count) And password = passwords(count) Then
            Open App.Path & "\gueststatus.txt" For Output As #2
                Print #2, "1"
                Close #2
                For count2 = 1 To 10000
                    frmsignon.top = frmsignon.top - 50
                    frmsignon.top = frmsignon.top
                Next count2
            frmsignon.Hide
            frmbuddylist.Show
            correct = 1
        End If
    Next count
        If correct <> 1 Then
            MsgBox "Ivalid Password/Username!", , "MiM:"
            txtuser.Text = ""
            txtpassword.Text = ""
            txtuser.SetFocus
        End If
        
End Sub

Private Sub lblclose_Click()
    End
End Sub
