VERSION 5.00
Begin VB.Form frmHomework 
   BackColor       =   &H00C0E0FF&
   Caption         =   "The Homework for Next Class is..."
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesigner 
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Text            =   "Designed by Amanda Aamodt"
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label lblDueDate 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Due Wednesday!"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label lblNumbers 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "#1-5, 8, 16, 19-24, 33-45(odd), 52"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblSection 
      BackColor       =   &H0080FF80&
      Caption         =   "Section 4"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblChapter 
      BackColor       =   &H0080FF80&
      Caption         =   "Chapter 3"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmHomework"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
