VERSION 5.00
Begin VB.Form frmAudience 
   BackColor       =   &H00800080&
   Caption         =   "Audience poll"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form4"
   ScaleHeight     =   4035
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Interval        =   8000
      Left            =   2760
      Top             =   1800
   End
   Begin VB.Timer Timer4 
      Interval        =   65
      Left            =   5160
      Top             =   3480
   End
   Begin VB.Timer Timer3 
      Interval        =   65
      Left            =   3720
      Top             =   3480
   End
   Begin VB.Timer Timer2 
      Interval        =   65
      Left            =   2160
      Top             =   3480
   End
   Begin VB.Timer Timer1 
      Interval        =   65
      Left            =   480
      Top             =   3600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "Pradeep de Noronha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblgraph4 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "64"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4920
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label lblgraph3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label lblgraph2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   5
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label lblgraph1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmaudience"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Who wants to be a Millionare.(millionare1.vbp)

'Form name: frmAudience; Form caption: Audience

'Author: Pradeep de Noronha

'Date written: 15th March, 2006

'Form Objective: This form is designed to show the audiences response to Question 1.
'                If the user decides to ask the audience for assistance on the first
'                question a graph will be drawn showing the audiences' opinion. This is
'                accomplished with the use of the Move button and Timers. The form
'                automatically hides itself after Timer 5 ends.

Private Sub Timer1_Timer()
    
    If lblgraph1.Top > 3000 Then
        lblgraph1.Move lblgraph1.Left - 0, lblgraph1.Top - 75
    End If
    
End Sub

Private Sub Timer2_Timer()
    
    If lblgraph2.Top > 3120 Then
        lblgraph2.Move lblgraph2.Left - 0, lblgraph2.Top - 75
    End If
    
End Sub

Private Sub Timer3_Timer()

    If lblgraph3.Top > 2880 Then
        lblgraph3.Move lblgraph3.Left - 0, lblgraph3.Top - 75
    End If
    
End Sub

Private Sub Timer4_Timer()

    If lblgraph4.Top > 1440 Then
        lblgraph4.Move lblgraph4.Left - 0, lblgraph4.Top - 75
    End If
    
End Sub

Private Sub Timer5_Timer()
    
    frmaudience.Hide
    
End Sub
