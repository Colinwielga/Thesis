VERSION 5.00
Begin VB.Form FrmMp3Format 
   BackColor       =   &H00000000&
   Caption         =   "Mp3 Format Application"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults1 
      Height          =   1455
      Left            =   1680
      ScaleHeight     =   1395
      ScaleWidth      =   5835
      TabIndex        =   11
      Top             =   8160
      Width           =   5895
   End
   Begin VB.PictureBox picresults 
      Height          =   1455
      Left            =   1680
      ScaleHeight     =   1395
      ScaleWidth      =   5835
      TabIndex        =   10
      Top             =   6480
      Width           =   5895
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "&After Picture"
      Height          =   855
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Before Picture"
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10080
      Width           =   1935
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      Caption         =   "AFTER"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Caption         =   "BEFORE"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label lbAuthor 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Author: Jeremy Phillips"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      Caption         =   "MP3 File Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "2.Click on the Before and After Picture buttons to see how the formatting takes your file names and removes the extra characters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   5520
      Width           =   8655
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "1. Click ""Start"" Button to format your files by deleting unwanted                 characters and symbols."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Width           =   8655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Directions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"VBPoject1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   8475
   End
End
Attribute VB_Name = "FrmMp3Format"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This Project Name is "Mp3 Format"
' Form Name is (frmMp3Format)
' Author "Jeremy Phillips"
' Date Oct 27th 2003
' Purpose : This Project is to help end user change their *.mp3
'           files to a standard format ersasing &, %, $ etc....
Option Explicit
Public Path As String
Dim Path1 As String

' Activates click function to get Format underway
Private Sub Command1_click()
    If Command1.Caption = "Start" Then
        Rename
    Else
        Command1.Caption = "Start"
    End If
End Sub
Private Sub Rename()
    Dim NewFileName As String
    Dim OldFileName As String
    Dim Path As String
    Dim CTR As Integer
    'Reads the file from the designated folder and reads the format inputed into NewFileName
    'make sure the path is the entire path all the way to the folder name with a backslash at the end, otherwise
    'The dir won't replace properly, and the program will crash.
    'I didn't place any MP3's in the folder because of their large size. I couldn't recall of the top of my head how much room our
    'M drive gives us. I figured that you would have a few you could place into a folder called "MP3" and run it.
    'If you need some, I will find a way to get you some.
    Path = "N:\CS130\Handin\Phillips_Jeremy\MP3\"
    NewFileName = Dir(Path & "*.mp3")
    
    'Do until loop that reads until the end of the file, and replaces certain characters with a null character
    Do Until NewFileName = ""
        OldFileName = NewFileName
        NewFileName = Replace(NewFileName, "$", "")
        NewFileName = Replace(NewFileName, ",", "")
        NewFileName = Replace(NewFileName, "!", "")
        NewFileName = Replace(NewFileName, "1", "")
        NewFileName = Replace(NewFileName, "2", "")
        NewFileName = Replace(NewFileName, "4", "")
        NewFileName = Replace(NewFileName, "5", "")
        NewFileName = Replace(NewFileName, "6", "")
        NewFileName = Replace(NewFileName, "7", "")
        NewFileName = Replace(NewFileName, "8", "")
        NewFileName = Replace(NewFileName, "9", "")
        NewFileName = Replace(NewFileName, "0", "")
        NewFileName = Replace(NewFileName, "@", "")
        NewFileName = Replace(NewFileName, "#", "")
        NewFileName = Replace(NewFileName, "%", "")
        NewFileName = Replace(NewFileName, "^", "")
        NewFileName = Replace(NewFileName, "&", "")
        NewFileName = Replace(NewFileName, "*", "")
        NewFileName = Replace(NewFileName, "(", "")
        NewFileName = Replace(NewFileName, ")", "")
        NewFileName = Replace(NewFileName, "?", "")
        NewFileName = Replace(NewFileName, "\", "")
        NewFileName = Replace(NewFileName, "+", "")
        NewFileName = Replace(NewFileName, ">", "")
        NewFileName = Replace(NewFileName, "<", "")
        NewFileName = Replace(NewFileName, "/", "")
        NewFileName = Replace(NewFileName, "`", "")
        NewFileName = Replace(NewFileName, "~", "")
        NewFileName = Replace(NewFileName, "=", "")
        NewFileName = Replace(NewFileName, "_", "")
        NewFileName = Replace(NewFileName, "{", "")
        NewFileName = Replace(NewFileName, "}", "")
        'Message box that asks if you like and accept the changes
        CTR = MsgBox("Do you want to change " & OldFileName & " to " & NewFileName & "?", vbYesNo)
        'enters the new name and file under same directory => thus replaces the file name
        If CTR = 6 Then
            Name Path & OldFileName As Path & NewFileName
        End If
        
        NewFileName = Dir
    Loop
End Sub
Private Sub cmdquit_Click()
End
End Sub

Private Sub Command2_Click()
Path1 = "N:\CS130\Handin\Phillips_Jeremy\"
picresults.Cls
picresults.Picture = LoadPicture(Path1 & "Before.jpg")
End Sub

Private Sub Command3_Click()
Path1 = "N:\CS130\Handin\Phillips_Jeremy\"
picresults.Cls
picresults1.Picture = LoadPicture(Path1 & "After.jpg")
End Sub
