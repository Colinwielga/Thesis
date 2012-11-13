VERSION 5.00
Begin VB.Form frmResult 
   BackColor       =   &H000000FF&
   Caption         =   "RESULTS"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhome 
      Caption         =   "GO HOME!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   600
      Picture         =   "frmResult.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   4755
      TabIndex        =   12
      Top             =   1080
      Width           =   4815
   End
   Begin VB.CommandButton cmdshowimage 
      Caption         =   "SHOW MY IMAGE!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   3495
   End
   Begin VB.PictureBox redblue 
      Height          =   7695
      Left            =   4935
      Picture         =   "frmResult.frx":0E39
      ScaleHeight     =   7635
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox redblack 
      Height          =   7575
      Left            =   4920
      Picture         =   "frmResult.frx":3BA23
      ScaleHeight     =   7515
      ScaleWidth      =   2475
      TabIndex        =   9
      Top             =   300
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox red 
      Height          =   7335
      Left            =   4800
      Picture         =   "frmResult.frx":790D1
      ScaleHeight     =   7275
      ScaleWidth      =   2835
      TabIndex        =   8
      Top             =   300
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox rurple 
      Height          =   7455
      Left            =   4800
      Picture         =   "frmResult.frx":BD137
      ScaleHeight     =   7395
      ScaleWidth      =   2595
      TabIndex        =   7
      Top             =   300
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox purple 
      Height          =   7335
      Left            =   5400
      Picture         =   "frmResult.frx":FB179
      ScaleHeight     =   7275
      ScaleWidth      =   2595
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox purplered 
      Height          =   7335
      Left            =   5520
      Picture         =   "frmResult.frx":154DD3
      ScaleHeight     =   7275
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox purpleblue 
      Height          =   7335
      Left            =   5400
      Picture         =   "frmResult.frx":19E4B1
      ScaleHeight     =   7275
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox purpleblack 
      Height          =   7335
      Left            =   5280
      Picture         =   "frmResult.frx":1EE8C3
      ScaleHeight     =   7275
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   300
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox blackblue 
      Height          =   7575
      Left            =   5520
      Picture         =   "frmResult.frx":259185
      ScaleHeight     =   7515
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox blackred 
      Height          =   7215
      Left            =   5175
      Picture         =   "frmResult.frx":2C864F
      ScaleHeight     =   7155
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   300
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox black 
      Height          =   7455
      Left            =   4815
      Picture         =   "frmResult.frx":34EB15
      ScaleHeight     =   7395
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHome_Click()
frmResult.Hide
frmDoll.Show
End Sub

Private Sub cmdshowimage_Click()
'Displays an image which is based on the colors that the user input earlier

If btmcolor = 1 And shcolor = 1 Then
    black.Visible = True
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = False
ElseIf btmcolor = 1 And shcolor = 2 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = True
    blackblue.Visible = False
    blackred.Visible = False
ElseIf btmcolor = 1 And shcolor = 3 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = True
ElseIf btmcolor = 2 And shcolor = 1 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = True
    blackblue.Visible = False
    blackred.Visible = False
ElseIf btmcolor = 2 And shcolor = 2 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = True
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = False
ElseIf btmcolor = 2 And shcolor = 3 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = True
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = False
ElseIf btmcolor = 3 And shcolor = 1 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = True
ElseIf btmcolor = 3 And shcolor = 1 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = True
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = False
ElseIf btmcolor = 3 And shcolor = 2 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = True
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = False
ElseIf btmcolor = 3 And shcolor = 3 Then
    black.Visible = False
    red.Visible = True
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = False
ElseIf btmcolor = 4 And shcolor = 1 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = True
    blackred.Visible = False
ElseIf btmcolor = 4 And shcolor = 2 Then
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = False
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = True
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = False
Else: btmcolor = 4 And shcolor = 3
    black.Visible = False
    red.Visible = False
    purple.Visible = False
    redblue.Visible = True
    redblack.Visible = False
    rurple.Visible = False
    purplered.Visible = False
    purpleblue.Visible = False
    purpleblack.Visible = False
    blackblue.Visible = False
    blackred.Visible = False

End If
picResults.Print "I feel prettier than ever!  Thank you!"
End Sub

