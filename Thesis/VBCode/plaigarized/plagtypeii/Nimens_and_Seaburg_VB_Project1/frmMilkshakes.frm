VERSION 5.00
Begin VB.Form frmMilkshakes
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   11940
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16185
   LinkTopic       =   "Form1"
   ScaleHeight     =   11940
   ScaleWidth      =   16185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch
      BackColor       =   &H00FF8080&
      Caption         =   "Milkshake Flavors"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10320
      Width           =   2055
   End
   Begin VB.CommandButton cmdGoback
      BackColor       =   &H00FF00FF&
      Caption         =   "Go Back"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10320
      Width           =   1935
   End
   Begin VB.Label lblPrize
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Prize: Milkshakes are free and you get your name and picture on the wall!!!"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   6
      Top             =   9360
      Width           =   9135
   End
   Begin VB.Label lblRules
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Rules: One Gallon of mallt milkshakes must be finished under 30 minutes without leaving your seat or throwing up."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4560
      TabIndex        =   5
      Top             =   8280
      Width           =   9135
   End
   Begin VB.Label lblSearch
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Click below to see if they have your favorite milkshake!"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Image imgMilkshakes
      BorderStyle     =   1  'Fixed Single
      Height          =   6045
      Left            =   4560
      Picture         =   "frmMilkshakes.frx":0000
      Top             =   2160
      Width           =   9060
   End
   Begin VB.Label lblChallenge
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   """Malt Milkshake Challenge"""
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   5400
      TabIndex        =   1
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label lblStlouis
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "St. Louis"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmMilkshakes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Man vs. Food
'frmMilkshakes
'Ty Nimens and Josh Seaburg
'February 2010
'Inform and have a button that when clicked it asks you for your favorite milkshake and then it searches to see if the store has it

Private Sub asdf_Click()
    frmMilkshakes.Hide
    frmMap.Show
End Sub

Private Sub aaaa_Click()
    End
End Sub

Private Sub sss_Click()
Dim CTR As Long, Milkshakes(1 To 15) As String, Found As Boolean, K As Integer, Shake As String
' fills array
    CTR = 0
    Open App.Path & "\milkshakes.txt" For Input As #1

    Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Milkshakes(CTR)
    Loop
    Close #1

    'find your favorite milkshake in an inputbox and messagebox
    Shake = InputBox("Enter your favorite milkshake.", "Milkshake")
    K = 0
    Found = False

Do While ((Not Found) And (K < CTR))
        K = K + 1
    If Shake = Milkshakes(K) Then
        Found = True
    End If
Loop

If (Not Found) Then
    MsgBox "Milkshake not found", vbCritical
Else
    MsgBox (Shake & " is one of the milkshakes!")
End If



End Sub
