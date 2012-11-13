VERSION 5.00
Begin VB.Form frmcited 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2715
      ScaleWidth      =   6555
      TabIndex        =   3
      Top             =   360
      Width           =   6615
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to home page"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdcite 
      Caption         =   "Show citations"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "frmcited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdcite_Click()   'picture Results for all of the citations
picResults.Print "The clothing, clothing prices, roster and schedule are all courtesy of GoJohnnies.com"
picResults.Print
picResults.Print "The pictures of all the wrestlers are all courtesy of Scott Padrnos and Facebook."
picResults.Print
picResults.Print "Also all of our Computer Science knowledge courtesy of Professor Miller/Holey."
End Sub

Private Sub cmdQuit_Click()
    End 'ending the program
End Sub

Private Sub cmdreturn_Click()
frmcited.Hide 'hiding the citations page
frmHome.Show 'showing the home page
End Sub
