VERSION 5.00
Begin VB.Form frmAdditionalServices 
   Caption         =   "Addtional Services"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdManicures 
      Caption         =   "Manicures"
      Height          =   1095
      Left            =   3240
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdFacials 
      Caption         =   "Facials"
      Height          =   1095
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdWomensHaircut 
      Caption         =   "Women's Haircut"
      Height          =   1215
      Left            =   3240
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdMensHaircut 
      Caption         =   "Men's Haircut"
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.PictureBox pbxResultsBeautySalon 
      Height          =   3015
      Left            =   6600
      ScaleHeight     =   2955
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblDanceLessons 
      Alignment       =   2  'Center
      Caption         =   "Dance Lessons"
      Height          =   1215
      Left            =   1320
      TabIndex        =   6
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label lblBeautySalon 
      Alignment       =   2  'Center
      Caption         =   "Beauty Salon"
      Height          =   1215
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
End
Attribute VB_Name = "frmAdditionalServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
