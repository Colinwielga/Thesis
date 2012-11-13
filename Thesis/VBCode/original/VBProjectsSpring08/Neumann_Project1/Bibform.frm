VERSION 5.00
Begin VB.Form bibform 
   Caption         =   "http://www.uksearchindex.com/tenpin-bowling.jpg"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   14565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "back home"
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
      Left            =   480
      TabIndex        =   6
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
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
      Left            =   11640
      TabIndex        =   5
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "http://www.irtc.org/ftp/pub/stills/2004-10-31/bowling.jpg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   3120
      Width           =   6855
   End
   Begin VB.Label Label4 
      Caption         =   "http://www.irtc.org/ftp/pub/stills/2005-12-31/bowling.jpg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Label Label3 
      Caption         =   "http://wallpapers.dpics.org/54__Bowling.htm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "http://www.uksearchindex.com/tenpin-bowling.jpg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1320
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "http://donnarobertsproshop.com/pin.gif"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   6975
   End
End
Attribute VB_Name = "bibform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdback_Click()
bibform.Hide
startform.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub

