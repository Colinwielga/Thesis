VERSION 5.00
Begin VB.Form frmAdScreen 
   BackColor       =   &H0000FFFF&
   Caption         =   "Purschase an Advertisement"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Stand"
      Height          =   735
      Left            =   7080
      TabIndex        =   9
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdAd8 
      BackColor       =   &H00C0C000&
      Caption         =   "Television campaign: $51.00"
      Height          =   495
      Left            =   2880
      MaskColor       =   &H000000FF&
      TabIndex        =   8
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdAd7 
      BackColor       =   &H00C0C000&
      Caption         =   "Radio campaign: $31.00"
      Height          =   495
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   7
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdAd6 
      BackColor       =   &H00C0C000&
      Caption         =   "Start a website: $21.00"
      Height          =   495
      Left            =   5640
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdAd3 
      BackColor       =   &H00C0C000&
      Caption         =   "Neighborhood posters: $6.00"
      Height          =   495
      Left            =   5640
      MaskColor       =   &H000000FF&
      TabIndex        =   5
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdAd5 
      BackColor       =   &H00C0C000&
      Caption         =   "Citywide posters: $13.00"
      Height          =   495
      Left            =   2880
      MaskColor       =   &H000000FF&
      TabIndex        =   4
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdAd2 
      BackColor       =   &H00C0C000&
      Caption         =   "Recruit friends: $5.00"
      Height          =   495
      Left            =   2880
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdAd4 
      BackColor       =   &H00C0C000&
      Caption         =   "Newspaper Ad: $9.00"
      Height          =   495
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdAd1 
      BackColor       =   &H00C0C000&
      Caption         =   "Recruit sister: $1.00"
      Height          =   495
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.Label lblAdInstruct 
      BackColor       =   &H00FFFF00&
      Caption         =   $"frmAdScreen.frx":0000
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmAdScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAd1_Click()
    If Cash >= 1 Then
        Fame = Fame + 1
        Cash = Cash - 1
    Else
        MsgBox "Not enough cash!"
    End If
End Sub

Private Sub cmdAd2_Click()
    If Cash >= 2 Then
        Fame = Fame + 2.1
        Cash = Cash - 2
    Else
        MsgBox "Not enough cash!"
    End If
End Sub

Private Sub cmdAd3_Click()
    If Cash >= 2.5 Then
        Fame = Fame + 2.8
        Cash = Cash - 2.5
    Else
        MsgBox "Not enough cash!"
    End If
End Sub

Private Sub cmdAd4_Click()
    If Cash >= 4 Then
        Fame = Fame + 4.6
        Cash = Cash - 4
    Else
        MsgBox "Not enough cash!"
    End If
End Sub

Private Sub cmdAd5_Click()
    If Cash >= 6 Then
        Fame = Fame + 6.9
        Cash = Cash - 6
    Else
        MsgBox "Not enough cash!"
    End If
End Sub

Private Sub cmdAd6_Click()
    If Cash >= 10 Then
        Fame = Fame + 13
        Cash = Cash - 10
    Else
        MsgBox "Not enough cash!"
    End If
End Sub

Private Sub cmdAd7_Click()
    If Cash >= 15 Then
        Fame = Fame + 18
        Cash = Cash - 15
    Else
        MsgBox "Not enough cash!"
    End If
End Sub

Private Sub cmdAd8_Click()
    If Cash >= 25 Then
        Fame = Fame + 35
        Cash = Cash - 25
    Else
        MsgBox "Not enough cash!"
    End If
End Sub

Private Sub cmdReturn_Click()
    frmAdScreen.Hide
    frmMainScreen.Show
End Sub
