VERSION 5.00
Begin VB.Form frmtalkto 
   BackColor       =   &H008080FF&
   Caption         =   "Who do you want to talk to?"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go back to the Boobery welcome page"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   3615
   End
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leave and continue on your tour de st. joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   3615
   End
   Begin VB.CommandButton cmdwhit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Whitney"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdtessie 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tessie"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdmara 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mara"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdbrooke 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Brooke"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdanna 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anna"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdmajor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort by major"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdfrom 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort by hometown"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmddrink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort by what they drink"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdname 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to see who you can talk to"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   4080
      ScaleHeight     =   4635
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   480
      Width           =   6375
   End
End
Attribute VB_Name = "frmtalkto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nameArray(1 To 5)
Dim boozeArray(1 To 5)
Dim townArray(1 To 5)
Dim majorArray(1 To 5)
Dim CTR As Integer


    'Project name:  Tour De St. Joe
    'Form:  frmtalkto, "Who do you want to talk to"
    'Author:  Brooke
    'Date:  3/10/08
    'Objective: To show who you could be talking to:  Using arrays to
    '          sort the people by different categories.
    
    
Private Sub cmdanna_Click()

    frmboobanna.Show
    frmtalkto.Hide

End Sub

Private Sub cmdback_Click()

    frmboob.Show
    frmtalkto.Hide


End Sub

Private Sub cmdbrooke_Click()

    frmboobbrooke.Show
    frmtalkto.Hide

End Sub

Private Sub cmddrink_Click()
Dim correctname As String
Dim correctdrink As String
Dim correctmajor As String
Dim correcttown As String
Dim Pass As Integer
Dim Pos As Integer
Dim X As Integer


    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
             If boozeArray(Pos) > boozeArray(Pos + 1) Then
                correctdrink = boozeArray(Pos)
                boozeArray(Pos) = boozeArray(Pos + 1)
                boozeArray(Pos + 1) = correctdrink
                correctname = nameArray(Pos)
                nameArray(Pos) = nameArray(Pos + 1)
                nameArray(Pos + 1) = correctname
                correctmajor = majorArray(Pos)
                majorArray(Pos) = majorArray(Pos + 1)
                majorArray(Pos + 1) = correctmajor
                correcttown = townArray(Pos)
                townArray(Pos) = townArray(Pos + 1)
                townArray(Pos + 1) = correcttown
            End If
        Next Pos
    Next Pass

    picresults.Cls

    For X = 1 To 5
        picresults.Print boozeArray(X); Tab(15); nameArray(X)
    Next X



End Sub

Private Sub cmdfrom_Click()

Dim correctbooze As String
Dim correctmajor As String
Dim correctname As String
Dim correcttown As String
Dim Pass As Integer
Dim Pos As Integer
Dim X As Integer


    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
             If townArray(Pos) > townArray(Pos + 1) Then
                correcttown = townArray(Pos)
                townArray(Pos) = townArray(Pos + 1)
                townArray(Pos + 1) = correcttown
                correctname = nameArray(Pos)
                nameArray(Pos) = nameArray(Pos + 1)
                nameArray(Pos + 1) = correctname
                correctmajor = majorArray(Pos)
                majorArray(Pos) = majorArray(Pos + 1)
                majorArray(Pos + 1) = correctmajor
                correctbooze = boozeArray(Pos)
                boozeArray(Pos) = boozeArray(Pos + 1)
                boozeArray(Pos + 1) = correctbooze
            End If
        Next Pos
    Next Pass

    picresults.Cls

    For X = 1 To 5
        picresults.Print townArray(X); Tab(15); nameArray(X)
    Next X

End Sub

Private Sub cmdleave_Click()

    frmjoetown.Show
    frmtalkto.Hide

End Sub

Private Sub cmdmajor_Click()

Dim correctbooze As String
Dim correcttown As String
Dim correctname As String
Dim correctmajor As String
Dim Pass As Integer
Dim Pos As Integer
Dim X As Integer


    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
             If majorArray(Pos) > majorArray(Pos + 1) Then
                correctmajor = majorArray(Pos)
                majorArray(Pos) = majorArray(Pos + 1)
                majorArray(Pos + 1) = correctmajor
                correctname = nameArray(Pos)
                nameArray(Pos) = nameArray(Pos + 1)
                nameArray(Pos + 1) = correctname
                correcttown = townArray(Pos)
                townArray(Pos) = townArray(Pos + 1)
                townArray(Pos + 1) = correcttown
                correctbooze = boozeArray(Pos)
                boozeArray(Pos) = boozeArray(Pos + 1)
                boozeArray(Pos + 1) = correctbooze
            End If
        Next Pos
    Next Pass

    picresults.Cls

    For X = 1 To 5
        picresults.Print majorArray(X); Tab(15); nameArray(X)
    Next X

End Sub

Private Sub cmdmara_Click()

    frmboobmara.Show
    frmtalkto.Hide

End Sub

Private Sub cmdname_Click()

Dim N As Integer

    'loading the arrays
        CTR = 0
            Open App.Path & "\boobgirls.txt" For Input As #1

            Do Until EOF(1)
                CTR = CTR + 1
                Input #1, nameArray(CTR), boozeArray(CTR), townArray(CTR), majorArray(CTR)
            Loop

        Close #1

    picresults.Cls

    'creating the parallel arrays
        For N = 1 To CTR
            picresults.Print nameArray(N)
        Next N

End Sub

Private Sub cmdtessie_Click()

    frmboobtessie.Show
    frmtalkto.Hide

End Sub

Private Sub cmdwhit_Click()

    frmboobwhit.Show
    frmtalkto.Hide

End Sub

Private Sub OLE1_Updated(Code As Integer)
    
    Open App.Path & "\07Track7.wma"

End Sub
