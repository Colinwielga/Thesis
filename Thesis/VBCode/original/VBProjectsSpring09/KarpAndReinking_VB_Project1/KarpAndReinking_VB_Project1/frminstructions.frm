VERSION 5.00
Begin VB.Form frminstructions 
   BackColor       =   &H80000012&
   Caption         =   "Instructions"
   ClientHeight    =   10380
   ClientLeft      =   2010
   ClientTop       =   645
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   Picture         =   "frminstructions.frx":0000
   ScaleHeight     =   10380
   ScaleWidth      =   12270
   Begin VB.CommandButton cmdplay 
      BackColor       =   &H0000C0C0&
      Caption         =   "Lets play Deal or No Deal!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      MaskColor       =   &H0000C0C0&
      TabIndex        =   1
      Top             =   9480
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"frminstructions.frx":1ED24
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   600
      TabIndex        =   9
      Top             =   8520
      Width           =   9495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "6. The above process will be repeated with 4, 3, 2 and 1 case(s)."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   8160
      Width           =   9495
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frminstructions.frx":1EE3A
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   7320
      Width           =   9495
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frminstructions.frx":1EEFC
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   6
      Top             =   6720
      Width           =   9495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Now Click on SIX more cases- the lower the amount of money in the case the higher your offer from the banker will be."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   6120
      Width           =   9495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Please click on the case that you believe to hold $1,000,000 (you will see if you're right at the end of the game)"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   5520
      Width           =   9495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Welcome to Deal or No Deal. The goal of this game is to win $1,000,000."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   5160
      Width           =   9495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   3240
      TabIndex        =   2
      Top             =   2520
      Width           =   15
   End
   Begin VB.Label lblinstructions 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Are you ready to play Deal or No Deal? Will you be our next big winner?  Here's how to play:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   7800
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frminstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project: Deal or No Deal
'frminstructiosn
'Holly Reinking and Danielle Karp
'Written 3/15/09
'Purpose: To give instructions on how to play the game for people who have never played before.

Private Sub cmdplay_Click()
frminstructions.Hide        'To hide one form and show another after the instructions are read
frmUsername.Show
End Sub


