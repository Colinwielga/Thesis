VERSION 5.00
Begin VB.Form frmbuddylist 
   Appearance      =   0  'Flat
   BackColor       =   &H00F3F3F3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Richard Motzko: Buddy List"
   ClientHeight    =   3135
   ClientLeft      =   11070
   ClientTop       =   630
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3990
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   2640
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   1800
   End
   Begin VB.Label lblguest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F3F3&
      Caption         =   "(Guest)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label lblsignoff 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "sign off"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label lblbuddylist 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "buddy list"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEB580&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmbuddylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub lblguest_Click()
    If lblguest.Enabled = True Then
        frmchat.Show
        c = 0
    End If
    If lblguest.Enabled = False Then
        c = 0
    End If
End Sub

Private Sub lblsignoff_Click()
    Dim i As Integer
    For i = 1 To 20000
        frmbuddylist.top = frmbuddylist.top - 50
        frmbuddylist.top = frmbuddylist.top
    Next i
    Open App.Path & "\rickstatus.txt" For Output As #1
        Print #1, "0"
    Close #1
    End
End Sub

Private Sub Timer1_Timer()
    If Timer1 = True Then
        Open App.Path & "\gueststatus.txt" For Input As #1
            Input #1, gueststatus
            Close #1
            If gueststatus = 1 Then
                lblguest.Enabled = True
                lblguest.ForeColor = &H80FF&
                lblguest.Caption = "Guest"
                buddy = "Guest"
            Else
                lblguest.Enabled = False
            End If
    End If
End Sub

Private Sub Timer2_Timer()
    Open App.Path & "\guestsaid.txt" For Input As #1
        Input #1, recieved
        Close #1
    If recieved <> clears Then
        frmchat.Show
    End If
End Sub
