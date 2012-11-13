VERSION 5.00
Begin VB.Form frmLake 
   BackColor       =   &H8000000D&
   Caption         =   "Bryan Mills"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   FillColor       =   &H00FFFF00&
   BeginProperty Font 
      Name            =   "Bernard MT Condensed"
      Size            =   72
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmLake.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFry 
      Caption         =   "Fish Fry"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6960
      TabIndex        =   9
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdHouse 
      BackColor       =   &H8000000D&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   8
      Top             =   2580
      Width           =   255
   End
   Begin VB.CommandButton cmdClub 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdReef 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton cmdPads 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   6480
      Width           =   375
   End
   Begin VB.CommandButton cmdRocks 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdPenninsula 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton cmdWeeds 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton cmdCattails 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblHotSpot 
      Caption         =   "Fishing Hot Spots"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmLake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program is a fishing guide service for White Bear Lake (final.vbp)
'this is the lake form (lake.frm)
'Bryan Mills
'March 24, 2006
'this is the form that tells you where to go and how to fish the area (lake.frm)
Option Explicit
    Dim found As Boolean
    Dim pos As Integer
Private Sub cmdCattails_Click()
    found = False
    pos = 0
    Do While found = False
        pos = pos + 1
        If location(pos) = 1 Then
            MsgBox "Location: " & method(pos), , "Location"
            MsgBox "This is how you will catch them: " & bait(pos), , "Bait"
            found = True
        End If
        Loop
        'this loops through the second array I used and prints the strings when the first value matches the search value
        found = True
End Sub

Private Sub cmdClub_Click()
    pos = 0
    found = False
    Do While found = False
        pos = pos + 1
        If location(pos) = 7 Then
            MsgBox "Location: " & method(pos), , "Location"
            MsgBox "This is how you will catch them: " & bait(pos), , "Bait"
            found = True
        End If
        Loop
        found = True
    End Sub

Private Sub cmdFry_Click()
    frmLake.Hide
    frmFry.Show
    
End Sub

Private Sub cmdHouse_Click()
    found = False
    pos = 0
    Do While found = False
        pos = pos + 1
        If location(pos) = 8 Then
            MsgBox "Location: " & method(pos), , "Location"
            MsgBox "This is how you will catch them: " & bait(pos), , "Bait"
            found = True
        End If
        Loop
        found = True
    End Sub

Private Sub cmdPads_Click()
    found = False
    pos = 0
    Do While found = False
        pos = pos + 1
        If location(pos) = 5 Then
            MsgBox "Location: " & method(pos), , "Location"
            MsgBox "This is how you will catch them: " & bait(pos), , "Bait"
            found = True
        End If
        Loop
        found = True
    End Sub

Private Sub cmdPenninsula_Click()
    found = False
    pos = 0
    Do While found = False
        pos = pos + 1
        If location(pos) = 3 Then
            MsgBox "Location: " & method(pos), , "Location"
            MsgBox "This is how you will catch them: " & bait(pos), , "Bait"
            found = True
        End If
        Loop
        found = True
    End Sub

Private Sub cmdReef_Click()
    found = False
    pos = 0
    Do While found = False
        pos = pos + 1
        If location(pos) = 3 Then
            MsgBox "Location: " & method(pos), , "Location"
            MsgBox "This is how you will catch them: " & bait(pos), , "Bait"
            found = True
        End If
        Loop
        found = True
    End Sub


Private Sub cmdRocks_Click()
    found = False
    pos = 0
    Do While found = False
        pos = pos + 1
        If location(pos) = 4 Then
            MsgBox "Location: " & method(pos), , "Location"
            MsgBox "This is how you will catch them: " & bait(pos), , "Bait"
            found = True
        End If
        Loop
        found = True
End Sub

Private Sub cmdWeeds_Click()
    found = False
    pos = 0
    Do While found = False
        pos = pos + 1
        If location(pos) = 2 Then
            MsgBox "Location: " & method(pos), , "Location"
            MsgBox "This is how you will catch them: " & bait(pos), , "Bait"
            found = True
        End If
        Loop
        found = True
    End Sub
