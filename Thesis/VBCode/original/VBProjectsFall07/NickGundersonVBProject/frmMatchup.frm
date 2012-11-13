VERSION 5.00
Begin VB.Form frmMatchup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Match-Up"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   FillColor       =   &H000000FF&
   FillStyle       =   5  'Downward Diagonal
   ForeColor       =   &H80000002&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   14
      Top             =   6840
      Width           =   2895
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Go On"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   13
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdBuss 
      Caption         =   "Buss Stop"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   12
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdDK 
      Caption         =   "The Special K Bars"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdInglis 
      Caption         =   "The Fresh Prince"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmdJason 
      Caption         =   "The Jon Kitna's"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   9
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdCase 
      Caption         =   "The Nimchucks"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmdAndy 
      Caption         =   "Make it Rain"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdGundy 
      Caption         =   "Vicks Pitbulls"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdPaul 
      Caption         =   "Torn ACL U"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmdPete 
      Caption         =   "PF Flyers"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdCubby 
      Caption         =   "The Cubby's"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdGervais 
      Caption         =   "Grrrrrr"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmdSchumacher 
      Caption         =   "Throw Some D's On It"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   10695
      Left            =   -4200
      Picture         =   "frmMatchup.frx":0000
      ScaleHeight     =   10635
      ScaleWidth      =   15435
      TabIndex        =   15
      Top             =   -2040
      Width           =   15495
      Begin VB.Label lblselect 
         Alignment       =   2  'Center
         Caption         =   "Please Select you're two teams"
         BeginProperty Font 
            Name            =   "Bodoni MT Black"
            Size            =   26.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6960
         TabIndex        =   16
         Top             =   2760
         Width           =   6975
      End
   End
   Begin VB.Label lblMatchup 
      Alignment       =   2  'Center
      Caption         =   "Please Choose the Matchup you would like to see for this fantasy Week"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmMatchup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR2 As Integer
'this page will have the user choose the two teams who play for the week
'once the team is selected the user cannot see the name
'a message box will appear everytime the user tries to select more then two teams


Private Sub cmdAndy_Click()
cmdAndy.Visible = False
If Team1 = 0 Then
    Team1 = 7
Else
    If Team2 = 0 Then
        Team2 = 7
    Else
        MsgBox ("Error already entered two teams")
    End If
End If
CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If

End Sub

Private Sub cmdBuss_Click()
cmdBuss.Visible = False
If Team1 = 0 Then
    Team1 = 12
Else
    If Team2 = 0 Then
        Team2 = 12
    Else
        MsgBox ("Error already entered two teams")
    End If
End If

CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If
End Sub

Private Sub cmdCase_Click()
cmdCase.Visible = False
If Team1 = 0 Then
    Team1 = 8
Else
    If Team2 = 0 Then
        Team2 = 8
    Else
        MsgBox ("Error already entered two teams")
    End If
End If
CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If
End Sub

Private Sub cmdCubby_Click()
cmdCubby.Visible = False
If Team1 = 0 Then
    Team1 = 3
Else
    If Team2 = 0 Then
        Team2 = 3
    Else
        MsgBox ("Error already entered two teams")
    End If
End If
CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If
End Sub

Private Sub cmdDK_Click()
cmdDK.Visible = False
If Team1 = 0 Then
    Team1 = 9
Else
    If Team2 = 0 Then
        Team2 = 9
    Else
        MsgBox ("Error already entered two teams")
    End If
End If

CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If

End Sub

Private Sub cmdGervais_Click()
cmdGervais.Visible = False
If Team1 = 0 Then
    Team1 = 2
Else
    If Team2 = 0 Then
        Team2 = 2
    Else
        MsgBox ("Error already entered two teams")
    End If
End If

CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If

End Sub

Private Sub cmdGundy_Click()
cmdGundy.Visible = False
If Team1 = 0 Then
    Team1 = 6
Else
    If Team2 = 0 Then
        Team2 = 6
    Else
        MsgBox ("Error already entered two teams")
    End If
End If

CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If

End Sub

Private Sub cmdInglis_Click()
cmdInglis.Visible = False
If Team1 = 0 Then
    Team1 = 11
Else
    If Team2 = 0 Then
        Team2 = 11
    Else
        MsgBox ("Error already entered two teams")
    End If
End If

CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If

End Sub

Private Sub cmdJason_Click()
cmdJason.Visible = False
If Team1 = 0 Then
    Team1 = 10
Else
    If Team2 = 0 Then
        Team2 = 10
    Else
        MsgBox ("Error already entered two teams")
    End If
End If

CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If

End Sub

Private Sub cmdNext_Click()
'this subroutine will move onto the next frm
'it also will check to make sure two teams have been selected

If Team1 = 0 Then
    MsgBox ("Please select 2 teams")
Else
    If Team2 = 0 Then
        MsgBox ("Please Select 1 more team")
    Else
        frmMatchup.Visible = False
        frmGame.Visible = True
    End If
End If

End Sub

Private Sub cmdPaul_Click()
cmdPaul.Visible = False
If Team1 = 0 Then
    Team1 = 5
Else
    If Team2 = 0 Then
        Team2 = 5
    Else
        MsgBox ("Error already entered two teams")
    End If
End If

CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If

End Sub

Private Sub cmdPete_Click()
cmdPete.Visible = False
If Team1 = 0 Then
    Team1 = 4
Else
    If Team2 = 0 Then
        Team2 = 4
    Else
        MsgBox ("Error already entered two teams")
    End If
End If

CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If

End Sub

Private Sub cmdQuit_Click()
End

End Sub

Private Sub cmdSchumacher_Click()

cmdSchumacher.Visible = False
If Team1 = 0 Then
    Team1 = 1
Else
    If Team2 = 0 Then
        Team2 = 1
    Else
        MsgBox ("Error already entered two teams")
    End If
End If

CTR2 = CTR2 + 1
If CTR2 = 2 Then
    cmdNext.Enabled = True
End If

End Sub

