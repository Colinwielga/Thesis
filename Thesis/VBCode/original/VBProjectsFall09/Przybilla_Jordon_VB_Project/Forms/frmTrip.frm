VERSION 5.00
Begin VB.Form frmTrip 
   Caption         =   "Hunting Trip"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   Picture         =   "frmTrip.frx":0000
   ScaleHeight     =   10785
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optUnguided 
      BackColor       =   &H00FFFF00&
      Caption         =   "Unguided"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   12
      Top             =   8640
      Width           =   1575
   End
   Begin VB.OptionButton optGuided 
      BackColor       =   &H00FFFF00&
      Caption         =   "Guided"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   11
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00008000&
      Caption         =   "Back to Home Page"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton cmdHunting 
      BackColor       =   &H00008000&
      Caption         =   "Back to Hunting Page"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuote 
      BackColor       =   &H00008000&
      Caption         =   "Get an Estimate!"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9480
      Width           =   3495
   End
   Begin VB.TextBox txtDeer 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Text            =   "0"
      Top             =   6960
      Width           =   3495
   End
   Begin VB.TextBox txtPeople 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Text            =   "0"
      Top             =   5280
      Width           =   3495
   End
   Begin VB.TextBox txtDays 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Text            =   "0"
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Label lblProcess 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Would you like a guided or unguided hunt?"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   6480
      Left            =   3960
      Picture         =   "frmTrip.frx":27997
      Top             =   3120
      Width           =   6480
   End
   Begin VB.Label lblDeer 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "How many deer total do you want to take?"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label lblPeople 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "How many people will be hunting?"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label lblDays 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "How many days would you like to hunt?"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Width           =   3015
   End
End
Attribute VB_Name = "frmTrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Terms'
'Authors: Jordon Przybilla'
'Date Written: October 15, 2009
'on this form the user will be allowed to set get an estimate for a whitetail hunt in minnesota
'this for will also display contact information for Whitetail X-treme Hunts

Option Explicit


Private Sub cmdHome_Click()
'returns the user to the home page
frmTrip.Hide
frmHome.Show

End Sub

Private Sub cmdHunting_Click()
'returns the user to the Hunting page
frmTrip.Hide
frmHunting.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdQuote_Click()
'this button will calculate a trip for the user based on what the user has put into the text boxes

Dim days As Integer, deer As Integer, people As Integer, guide As String, quote As Long

days = txtDays.Text
people = txtPeople.Text
deer = txtDeer.Text


If optGuided.Value = True Then
    quote = (days * 125) + (people * 200) + (deer * 200) + (250 * days)
    MsgBox "Your trip would cost approximately" & FormatCurrency(quote) & "."
ElseIf optUnguided.Value = True Then
    quote = (days * 125) + (people * 200) + (deer * 200)
    MsgBox "Your trip would cost approximately" & FormatCurrency(quote) & "."
End If



End Sub



