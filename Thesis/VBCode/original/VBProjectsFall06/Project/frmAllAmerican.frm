VERSION 5.00
Begin VB.Form frmAllAmerican 
   BackColor       =   &H00000080&
   Caption         =   "All-American Gophers"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFill 
      BackColor       =   &H0000FFFF&
      Caption         =   "View All Americans"
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H0000FFFF&
      Caption         =   "Home"
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000FFFF&
      Caption         =   "Back"
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort All-Americans by Name"
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FFFF&
      Height          =   6975
      Left            =   480
      ScaleHeight     =   6915
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label lblAllAmericans 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "First Team All-Americans"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   2640
   End
End
Attribute VB_Name = "frmAllAmerican"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmAllAmerican
'Cole and John
'10/30/06
'Objective: The objective of this form is to show the user past All-American hockey
'players at the University of Minnesota.  The user can view the All-Americans by
'year or by sorting alphabetically according to their last name.

Option Explicit

Dim Year(1 To 35) As Integer
Dim Pname(1 To 35) As String

Private Sub cmdBack_Click()
    frmAllAmerican.Visible = False
    frmHistory.Visible = True
End Sub

Private Sub cmdFill_Click()

Dim Pos As Integer

    picResults.Cls
    
    Open App.Path & "\AllAmerican.txt" For Input As #1      'opens text file
    Pos = 0
    
        Do Until Pos = 35                       'puts text file into array with 35 lines
            Pos = Pos + 1
            Input #1, Year(Pos), Pname(Pos)     'the first column in the text file is Year, second is the player name
        Loop
     Close #1
     
    For Pos = 1 To 35
        picResults.Print Year(Pos), Pname(Pos)
    Next Pos
    
End Sub


Private Sub cmdHome_Click()
    frmAllAmerican.Visible = False
    frmMain.Visible = True
End Sub
cmdSort_Click
Private Sub cmdSort_Click()
Dim pass As Integer
Dim temp1 As String
Dim comp As Integer
Dim Pos As Integer

    Pos = 0

    For pass = 1 To 34                              'makes n-1 passes through the list
        For comp = 1 To 35 - pass
            If Pname(comp) > Pname(comp + 1) Then    'on each pass, compares each name to its neighbor
                temp1 = Pname(comp)
                Pname(comp) = Pname(comp + 1)
                Pname(comp + 1) = temp1
                Pos = Year(comp)
                Year(comp) = Year(comp + 1)            'if out of alphabetic order, switches positiions
                Year(comp + 1) = Pos
            End If
        Next comp
    Next pass
    
    picResults.Cls
    
    For Pos = 1 To 35
        picResults.Print Year(Pos), Pname(Pos)      'prints the names in alphabetic order
    Next Pos
End Sub
