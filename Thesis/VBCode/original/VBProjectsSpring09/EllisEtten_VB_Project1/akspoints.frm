VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   FillColor       =   &H0080FF80&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcPoints 
      BackColor       =   &H0000C000&
      Caption         =   "Start Calcutating points!"
      Height          =   615
      Left            =   3600
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   3375
   End
   Begin VB.Image imgAKS 
      Height          =   7560
      Left            =   2520
      Picture         =   "aks points.frx":0000
      Top             =   120
      Width           =   5400
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalcPoints_Click()
    frmBigLittle.Show
    frmWelcome.Hide
End Sub
