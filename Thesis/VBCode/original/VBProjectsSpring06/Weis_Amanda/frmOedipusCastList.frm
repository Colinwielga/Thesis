VERSION 5.00
Begin VB.Form frmOedipusCastList 
   BackColor       =   &H00000000&
   Caption         =   "Oedipus Tex Cast List"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   Picture         =   "frmOedipusCastList.frx":0000
   ScaleHeight     =   8085
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      ToolTipText     =   "Press after typing in desired character to learn more about them."
      Top             =   6960
      Width           =   1335
   End
   Begin VB.PictureBox picCastList 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5895
      Left            =   9000
      ScaleHeight     =   5835
      ScaleWidth      =   5835
      TabIndex        =   4
      Top             =   240
      Width           =   5895
   End
   Begin VB.TextBox txtInfo 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   7080
      Width           =   3135
   End
   Begin VB.CommandButton cmdDisplayCastList 
      Caption         =   "Display Cast List"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   1
      ToolTipText     =   "Click to read the cast list."
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   0
      ToolTipText     =   "Click to go back to previous form."
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label lblBillie 
      BackColor       =   &H00000000&
      Caption         =   "The Greek Chorus mourns Billie Jo's Death"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label lblDesign 
      BackColor       =   &H00000000&
      Caption         =   "Design By Amanda Weis"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9120
      TabIndex        =   5
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Which character would you like to learn more about?"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   6240
      Width           =   2895
   End
End
Attribute VB_Name = "frmOedipusCastList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'declare all arrays and variables
    Dim Character(1 To 100) As String
    Dim Year(1 To 100) As String
    Dim Names(1 To 100) As String
    Dim Arraysize As Integer
    'open and read file and print cast list
Private Sub cmdDisplayCastList_Click()
    Dim Pos As Integer
    Pos = 0
    picCastList.Cls
    Open App.Path & "\OedipusTex.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Character(Pos), Names(Pos), Year(Pos)
        picCastList.Print Character(Pos); Tab(25); Names(Pos), Year(Pos)
    Loop
    Close #1
    Arraysize = Pos
End Sub
    'create button to go back to previous form
Private Sub cmdGoBack_Click()
    frmOedipusCastList.Hide
    frmOedipusTex.Show
End Sub
    'use case so that the user can input a character and learn more about them through a message box
Private Sub cmdSearch_Click()
    Select Case txtInfo.Text
        Case "Oedipus Tex"
            MsgBox "Oedipus Tex is the main role of the opera.  He unknowingly murders his father and marries his mother."
        Case "Billie Jo Casta"
            MsgBox "Billie Jo Casta is Oedipus Tex's mother and wife.  She kills herself when a fortune teller tells her the truth about her relationship with her husband."
        Case "Madame Peep"
            MsgBox "Madame Peep is the fortune teller in town who tells Oedipus and Billie Jo that the reason there's a plague hanging around town is because of their sinful behavior."
        Case "Narrator"
            MsgBox "The narrator is there to keep the audience up to par on what's happening within the opera.  He knows everything that's goin on."
        Case "Greek Chorus"
            MsgBox "The Greek Chorus takes pleasure in knowing the terrible 'tragedy' that has befallen Oedipus and also helps the narrator tell the audience what is going on."
        Case Else
            MsgBox "Check Spelling", vbCritical
    End Select
End Sub



