VERSION 5.00
Begin VB.Form frmDepartments 
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave Ben's Hockey Goods"
      Height          =   855
      Left            =   4440
      TabIndex        =   6
      Top             =   5760
      Width           =   2895
   End
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "Proceed to Checkout"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton cmdAccessories 
      Caption         =   "Accessories"
      Height          =   735
      Left            =   5280
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdSkates 
      Caption         =   "Skates"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdHelments 
      Caption         =   "Helments"
      Height          =   735
      Left            =   5280
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdSticks 
      Caption         =   "Sticks"
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdPadding 
      Caption         =   "Padding"
      Height          =   735
      Left            =   240
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmDepartments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
