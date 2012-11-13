VERSION 5.00
Begin VB.Form frmFDR 
   Caption         =   "Inaugural Speech Mad Lib -  Franklin D. Roosevelt, 1941"
   ClientHeight    =   7500
   ClientLeft      =   2295
   ClientTop       =   2115
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10875
   Begin VB.PictureBox picAmercianFDR 
      Height          =   7695
      Left            =   -360
      Picture         =   "frmFDR.frx":0000
      ScaleHeight     =   7635
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   -120
      Width           =   11415
      Begin VB.PictureBox picDisplay 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   720
         ScaleHeight     =   4995
         ScaleWidth      =   10035
         TabIndex        =   5
         Top             =   480
         Width           =   10095
      End
      Begin VB.CommandButton cmdRead 
         Caption         =   "Input Words"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   4
         Top             =   6120
         Width           =   2055
      End
      Begin VB.CommandButton cmdGoBack 
         Caption         =   "Go Back"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   2
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "Display Words"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         TabIndex        =   1
         Top             =   6120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmFDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As String, c1 As String, c2 As String, c3 As String, c4 As String, c5 As String, c6 As String
Dim c7 As String, c8 As String, c9 As String, c10 As String, c11 As String, c12 As String


Private Sub cmdDisplay_Click()
    picDisplay.Print "On each national day of inauguration since "; c; ", the people have renewed their sense of dedication to the United States."
    picDisplay.Print "In "; c1; "'s day the task of the people was to create and weld together a nation."
    picDisplay.Print "In Lincoln's day the task of the people was to "; c2; " that Nation from disruption from within."
    picDisplay.Print "In this day the task of the people is to "; c3; " that Nation and its institutions from disruption from without."
    picDisplay.Print "To us there has come a time, in the midst of "; c4; " happenings, to pause for a moment and take stock--to recall"
    picDisplay.Print "what our place in history has been, and to rediscover what we are and what we may be."
    picDisplay.Print "If we do not, we risk the real peril of inaction."
    picDisplay.Print
    picDisplay.Print "Lives of nations are determined not by the number of "; c5; ", but by the lifetime of the human spirit."
    picDisplay.Print "The life of a man is three-score years and ten: a little more, a little less."
    picDisplay.Print "The life of a nation is the fullness of the measure of its will to live."
    picDisplay.Print
    picDisplay.Print "There are "; c6; " who doubt this."
    picDisplay.Print "There are men who believe that "; c7; ", as a form of Government and a frame of life,"
    picDisplay.Print "is limited or measured by a kind of "; c8; " and "; c9; " fate that, for some unexplained reason,"
    picDisplay.Print "tyranny and slavery have become the surging wave of the future--and that freedom is an ebbing tide."
    picDisplay.Print
    picDisplay.Print "But we Americans know that this is not true."
    picDisplay.Print
    picDisplay.Print "Eight years ago, when the life of this Republic seemed frozen by a fatalistic terror, we proved that this is not true. "
    picDisplay.Print "We were in the midst of shock--but we acted."
    picDisplay.Print "We acted "; c10; ", "; c11; ", "; c12; "."


End Sub

Private Sub cmdGoBack_Click()
    frmBeginMadLib.Show
    frmFDR.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRead_Click()
    Open App.Path & "\wordFDR.txt" For Append As #1
    c = InputBox("Enter a year.")
    c1 = InputBox("Enter any President's last name.")
    c2 = InputBox("Enter a verb.")
    c3 = InputBox("Enter a verb.")
    c4 = InputBox("Enter an adjective.")
    c5 = InputBox("Enter a type of farm animal. (plural)")
    c6 = InputBox("Enter a noun. (plural)")
    c7 = InputBox("Enter a form of government.")
    c8 = InputBox("Enter an adjective.")
    c9 = InputBox("Enter an adjective.")
    c10 = InputBox("Enter an adjective ending with -ly.")
    c11 = InputBox("Enter an adjective ending with -ly.")
    c12 = InputBox("Enter an adjective ending with -ly.")
    Write #1, c, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12
    cmdDisplay.Enabled = True
    Close #1
    MsgBox ("Good Job, when you are ready to see your masterpiece, CLICK on Display Words.")
End Sub
