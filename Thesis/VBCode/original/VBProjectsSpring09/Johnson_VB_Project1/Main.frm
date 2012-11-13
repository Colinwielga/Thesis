VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C000&
   Caption         =   "America"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17895
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   17895
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   7500
      Left            =   3120
      ScaleHeight     =   3.024
      ScaleMode       =   0  'User
      ScaleWidth      =   8.323
      TabIndex        =   9
      Top             =   1440
      Width           =   11985
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit Program"
      Height          =   855
      Left            =   11280
      TabIndex        =   7
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton cmdGettysburg 
      Caption         =   "Gettysburg, PA"
      Height          =   735
      Left            =   15120
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdDublin 
      Caption         =   "Dublin (Hometown)"
      Height          =   735
      Left            =   1200
      TabIndex        =   5
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdLDAC 
      Caption         =   "Leadership Development and Assessment Course (LDAC)"
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdDCLT 
      Caption         =   "Drill Cadet Leadership Training (DCLT)"
      Height          =   735
      Left            =   15120
      TabIndex        =   3
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdAirborne 
      Caption         =   "US Army Airborne School"
      Height          =   735
      Left            =   15120
      TabIndex        =   2
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdHargrave 
      Caption         =   "Hargrave Military Academy"
      Height          =   735
      Left            =   15120
      TabIndex        =   1
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdVegas 
      Caption         =   "Vegas"
      Height          =   800
      Left            =   1200
      TabIndex        =   0
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblShoutOut 
      BackColor       =   &H0000FFFF&
      Caption         =   "MORE COOL STUFF HERE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   15120
      TabIndex        =   12
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblShoutOut 
      BackColor       =   &H0000FFFF&
      Caption         =   "COOL STUFF HERE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Tim Johnson   3/20    To give the user a useable interface for the information I wish to cover"
      Height          =   615
      Left            =   14520
      TabIndex        =   10
      Top             =   9840
      Width           =   2895
   End
   Begin VB.Label lblMainTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Where Has TJ Been?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   54
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      TabIndex        =   8
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR As Integer, Quotes(1 To 8) As String, mainpix(1 To 8) As String 'Declares all variables and arrays for this form

Private Sub cmdAirborne_Click()                                 'Goes to Airborne Form
CTR = 8                                                         'sets CTR for next few steps

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(CTR))    'gives a preview of what is to come

MsgBox Quotes(CTR), , "Did You Know?"                           'presents an interesting quote about the upcoming topic

frmMain.Hide                                                    'Goes to Airborne Form
frmAirborne.Show

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(1))      'brings back main picture for the users return

End Sub

Private Sub cmdDCLT_Click()                                     'Goes to DCLT Form
CTR = 7                                                         'sets CTR for next few steps

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(CTR))    'gives a preview of what is to come

MsgBox Quotes(CTR), , "Did You Know?"                           'presents an interesting quote about the upcoming topic

frmMain.Hide                                                    'Goes to DCLT Form
frmDCLT.Show

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(1))      'brings back main picture for the users return

End Sub

Private Sub cmdDublin_Click()                                   'Goes to Dublin Form
CTR = 4                                                         'sets CTR for next few steps

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(CTR))    'gives a preview of what is to come

MsgBox Quotes(CTR), , "Did You Know?"                           'presents an interesting quote about the upcoming topic

frmMain.Hide                                                    'Goes to DCLT Form
frmDublin.Show

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(1))      'brings back main picture for the users return

End Sub

Private Sub cmdGettysburg_Click()                               'Goes to Gettysburg Form
CTR = 5                                                         'sets CTR for next few steps

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(CTR))    'gives a preview of what is to come

MsgBox Quotes(CTR), , "Did You Know?"                           'presents an interesting quote about the upcoming topic

frmMain.Hide                                                    'Goes to Gettysburg Form
frmGettysburg.Show

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(1))      'brings back main picture for the users return

End Sub

Private Sub cmdHargrave_Click()                                 'Goes to Hargrave Form
CTR = 6                                                         'sets CTR for next few steps

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(CTR))    'gives a preview of what is to come

MsgBox Quotes(CTR), , "Did You Know?"                           'presents an interesting quote about the upcoming topic

frmMain.Hide                                                    'Goes to Hargrave Form
frmHargrave.Show

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(1))      'brings back main picture for the users return

End Sub

Private Sub cmdLDAC_Click()                                     'Goes to LDAC Form
CTR = 2                                                         'sets CTR for next few steps

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(CTR))    'gives a preview of what is to come

MsgBox Quotes(CTR), , "Did You Know?"                           'presents an interesting quote about the upcoming topic

frmMain.Hide                                                    'Goes to LDAC Form
frmLDAC.Show

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(1))      'brings back main picture for the users return

End Sub

Private Sub cmdQuit_Click()                                     'Ends Program where you are
    End                                                         'Ends Program where you are
End Sub

Private Sub cmdVegas_Click()                                    'Goes to Vegas Form
CTR = 3                                                         'sets CTR for next few steps

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(CTR))    'gives a preview of what is to come

MsgBox Quotes(CTR), , "Did You Know?"                           'presents an interesting quote about the upcoming topic

frmMain.Hide                                                    'Goes to LDAC Form
frmVegas.Show

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(1))      'brings back main picture for the users return

End Sub

Private Sub Form_Load()                                         'Loads all the pictures used on this form

Open App.Path & "\mainpix.txt" For Input As #2                  'opens text file for reading into an array

CTR = 0                                                         'sets CTR to 0 for the array

Do While Not EOF(2)                                             'reads pictures into arrays until the file has no more info
    CTR = CTR + 1
    Input #2, mainpix(CTR), Quotes(CTR)
Loop
Close #2                                                        'closes file

picMain.Picture = LoadPicture(App.Path & "\" & mainpix(1))      'loads the main pic the first time the user comes to this form

MsgBox Quotes(1), , "Did You Know?"                             'gives a pop up factoid about America

End Sub

